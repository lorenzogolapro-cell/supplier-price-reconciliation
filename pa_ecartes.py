# -*- coding: utf-8 -*-
"""
Purchase-price workstream: what we set aside, in separate lists and
nothing else.

    python pa_ecartes.py

WHY SEVERAL SHEETS, AND NOT ONE
    These are decisions of different natures, and mixing them up would
    lose the information:

    "Écarté"             the product exists and sells; it is WE who are
                         postponing it (beds, wheelchairs, lift chairs,
                         VPH). A scope decision, reversible.

    "Fournisseur arrêté" we no longer buy from THEM; there is nothing
                         wrong with the product itself. A supply finding,
                         dated (`fournisseurs_arretes.py`), not a guess.

    "Arrêt fabricant"    the product cannot be found in the supplier's
                         catalogue; it has probably been discontinued. A
                         finding about the outside world, not an internal
                         decision.

    An item set aside may come back; a dropped supplier or a
    discontinued product will not. Hence the separate sheets.

WHAT FILLS EACH SHEET
    Écarté             the annotation written by hand in
                       pa_decisions.xlsx, PLUS the whole heavy-equipment
                       family of the scope (FAMILLES), PLUS the suppliers
                       set aside by management.
    Fournisseur arrêté `fournisseurs_arretes.est_arrete()`, checked on the
                       supplier AND on the manufacturer, otherwise an item
                       for which the WMS only fills in the manufacturer
                       escapes the filter. That happened on 18/09: item
                       48553 (Supplier T, empty supplier field,
                       manufacturer "FRST FRANCE SAS") had been given a
                       real 2026 price despite ordering having been found
                       to have stopped on 15/09.
    Arrêt fabricant    the hand-written annotation, and that ALONE.
                       Nothing here is inferred: "discontinued" is a claim
                       about a product, it is not something to guess.

    The WMS neither dates nor flags the discontinuation of a PRODUCT (no
    Type 79 in the scope as of 17/09): there is therefore no automatic
    corroboration for "Arrêt fabricant". The "Votre annotation" column is
    the evidence, and the only one. The dropping of a SUPPLIER, by
    contrast, is written down and dated elsewhere: not the same source,
    nor the same level of certainty.

THE MISSING SUPPLIER
    423 items of the scope have no supplier. 420 have a manufacturer in
    the WMS extract, and that is what we fall back on; the "Origine du
    fournisseur" column always says where the value comes from, so that a
    known supplier is never confused with an assumed manufacturer. The
    remaining 3 are named in the console output.

NOTHING IS WRITTEN ANYWHERE ELSE
    This module reads and sorts. It touches neither the price owner's
    workbook, nor the prices, nor the scope.
"""

from __future__ import annotations

import re
import sys
import unicodedata
from datetime import date
from functools import lru_cache
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import warnings  # noqa: E402

warnings.filterwarnings("ignore")

import pandas as pd  # noqa: E402

import fournisseurs_arretes  # noqa: E402
import mise_en_forme  # noqa: E402
import pa_annotations_livrable  # noqa: E402
from extracteurs.base import chemin_lisible  # noqa: E402
from main import SORTIE_ACHATS  # noqa: E402

ETAT = SORTIE_ACHATS / "pa_etat_par_article.xlsx"
DECISIONS = SORTIE_ACHATS / "pa_decisions.xlsx"
SORTIE = SORTIE_ACHATS / "pa_ecartes_et_arrets.xlsx"

# The heavy-equipment families, word for word as the WMS spells them.
# Management decision on 17/09: beds, wheelchairs, lift chairs and VPH
# leave the current workstream, too many options, too many variants, one
# price per configuration. We file them away, we do not throw them out.
#
# "Transfert" is NOT in here: transfer handles are priced normally and
# some have already been matched (see item 31128).
FAMILLES = (
    "Fauteuil Roulant",
    "FAUTEUIL ROULANT MANUEL",
    "Accessoires pour fauteuils roulants",
    "Faut releveurs et repos",
    "Lit",
    "Scooters  tricycles poussettes",
    "Elevateur de bain",
    "LES APPAREILS DE VERTICALISATION",
)

# Set aside by management. We look at the MANUFACTURER as much as the
# supplier: 67 items carry Supplier H as manufacturer without carrying it
# as supplier, and were slipping through the filter.
MIS_DE_COTE = re.compile(r"FRSB|FRSH|FRSR", re.I)

# What your words mean. Discontinuation wins: it is more precise than a
# family-level exclusion, and it comes from a finding in the catalogue.
#
# "inexistant" is typed by hand, so rarely twice the same way: we read
# "innexistant" AND "inenxistant" on the same day. Hence the three-letter
# gap between "in" and "xist"; an annotation lost to a typo means an item
# we still believe is sellable.
DIT_ARRET = re.compile(
    r"arr[eêé]t|n.existe plus|\bin[a-z]{0,3}xist|introuvable", re.I)
DIT_ECARTE = re.compile(r"[eé]cart|releveur|\bvph\b|exclu", re.I)

# NO CATALOGUE: the supplier never sent us a price list (18/09)
#
# A fourth category, and it does not say the same thing as the other
# three. "Écarté" is a choice, "arrêté" is a finding about the product or
# the relationship; here there is NEITHER a choice NOR a stop: the raw
# material is simply missing. The product sells, the supplier exists, we
# have nothing to read.
#
# Mixing the two distorts the reading in both directions: it inflates the
# work remaining with items on which no work at all is possible, and it
# hides the one action that would unblock them, asking for the price list.
#
# HOW THE CATEGORY FILLS UP: two routes, and the first is almost always
# enough:
#
#   1. DERIVED from the "aucun tarif reçu de ce fournisseur" reason that
#      `pa_etat_par_article` already sets. 56 suppliers, 313 active items
#      as of 18/09. No list to maintain: a supplier who finally sends
#      their price list leaves the category on its own, at the next run.
#
#   2. NAMED by hand below, for what the reason cannot catch, first and
#      foremost the items with NO supplier in the WMS, which carry a
#      different reason ("aucun fournisseur au périmètre") and for which
#      only the manufacturer is known: Supplier FA is the textbook case,
#      its 12 items have no supplier at all.
#
# List dictated by the business on 18/09. We match on the supplier AND on
# the manufacturer, for the reason above.
#
# DELIBERATELY NOT IN THIS LIST:
#   Suppliers Q and Q2: we have their catalogue
#       (`catalogues/FRSQ-FRSQ2/2026/`), they are two names for one and
#       the same supplier. If their items come out without a price, that
#       is a matching defect, not a missing price list; see the note in
#       `classer()`.
#   Supplier FP: taken over by Supplier H, so it is a DROPPED supplier,
#       there will never be a catalogue
#       (`fournisseurs_arretes.ARRETES`).
SANS_CATALOGUE = (
    "FRSFA",
    "FRSFB MEDICAL",
    "FRSFC",
    "FRSFD",
    "FRSFE",
    "FRSFF",
    "FRSBS",
    "FRSFG",
    "FRSBN",
    "FRSFH",
    "FRSFI",              # "SAS FRSFI - FRSFI BIS": the two legal names
    "FRSFI BIS",          #   designate the same supplier
    "FRSFJ CENTRE",       # the acronym alone would catch two other
                          #   suppliers whose name contains it
    "FRSFK",
    "FRSFL",
    "FRSFM NORD",         # the acronym alone is too short for a whole word
    "FRSFN-SUFFIXE",
    "FRSFN SUFFIXE",      # two spellings, hyphen or space
    "FRSFO",
)

_SANS_CATALOGUE = re.compile(
    "|".join(re.escape(nom) for nom in SANS_CATALOGUE), re.I)

# BRANDS THAT IDENTIFY A SUPPLIER INSIDE A LABEL (18/09)
#
# Rule laid down by the business: "if a supplier's name is in the label,
# you put it in discontinued". The case that prompted it:
#
#   46675  STETHOSCOPE ... MODELE FRSBD GRIS  -> Supplier A,  no price
#   46677  STETHOSCOPE ... MODELE FRSBD VERT  -> Supplier BD, has a price
#
# The same product twice: the one that is correctly attached is priced,
# the other is not. The item is not untraceable, it is WRONGLY ATTACHED
# in the WMS, and we do not fix the WMS from this repository.
#
# A CURATED LIST, NOT A DERIVED ONE, and that is deliberate. Deriving the
# brands from the 127 supplier names fired 14 times, FOUR of them on the
# word "CONFORT": "OPTION ASSISE CONFORT ROLLATOR" has nothing to do with
# the legal name of Supplier CT nor with that of Supplier BV, both of
# which contain that word. It is the same trap as `MOTS_TROP_COURANTS`
# elsewhere in the workstream: an everyday French word cannot serve as an
# identifier.
#
# Only add a name here that means NOTHING other than a brand.
MARQUES_FOURNISSEUR = (
    "FRSBD", "FRSV", "FRSF", "FRSG", "FRSC", "FRSAU",
    "FRSAE", "FRSAK", "FRSAF", "FRSD", "FRSX", "FRSO",
    "FRSE", "FRSB", "FRSH", "FRSL", "FRSBF",
    "FRSBG", "FRSFQ", "FRSW", "FRSAI",
)


def _plein(serie: pd.Series) -> pd.Series:
    return serie.fillna("").astype(str).str.strip().replace("nan", "")


def _aplati(texte) -> str:
    """A name reduced to its letters and digits, so that comparison is not
    trapped by punctuation ("FRS.AE MEDICAL" vs "FRSAE")."""
    sans_accent = unicodedata.normalize("NFKD", str(texte or "").upper())
    sans_accent = "".join(c for c in sans_accent
                          if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]", "", sans_accent)


@lru_cache(maxsize=1)
def _catalogues_connus() -> tuple:
    """The names under which a catalogue exists: folders + registry.

    Serves to answer ONE question: "do we have anything to read for this
    supplier?". Two sources because they do not overlap: a folder can
    exist with no registry entry (nobody reads it yet), and a registry
    pattern can point at a folder with a different name.
    """
    from main import CATALOGUES, FOURNISSEURS

    noms = set()
    if CATALOGUES.is_dir():
        for dossier in CATALOGUES.iterdir():
            if dossier.is_dir() and dossier.name != "Archives":
                # "FRSQ-FRSQ2" stands for FRSQ AND for FRSQ2: we split on
                # the separators, otherwise FRSQ2 would pass for having no
                # catalogue when its price list is in that very folder.
                for morceau in re.split(r"[-+&/]", dossier.name):
                    if len(_aplati(morceau)) >= 4:
                        noms.add(_aplati(morceau))
    for entree in FOURNISSEURS:
        if len(_aplati(entree["motif_wms"])) >= 4:
            noms.add(_aplati(entree["motif_wms"]))
    return tuple(sorted(noms))


def a_un_catalogue(nom) -> bool:
    """Is there a catalogue for this supplier or this manufacturer?"""
    p = _aplati(nom)
    if len(p) < 4:
        return False
    return any(c in p or p in c for c in _catalogues_connus())


def perimetre() -> pd.DataFrame:
    d = pd.read_excel(chemin_lisible(ETAT), skiprows=3, dtype=str)
    for c in d.columns:
        d[c] = _plein(d[c])
    d["Code article"] = d["Code article"].str.replace(r"\.0$", "", regex=True)
    d["Vendu 12 mois"] = d["VENTE_12M"].str.lower().map(
        {"true": "oui"}).fillna("")

    # The missing supplier is filled in from the manufacturer, and that
    # is stated.
    manque = d["Fournisseur"] == ""
    d["Origine du fournisseur"] = "WMS"
    d.loc[manque, "Origine du fournisseur"] = "introuvable"
    repris = manque & (d["Nom fabriquant"] != "")
    d.loc[repris, "Fournisseur"] = d.loc[repris, "Nom fabriquant"]
    d.loc[repris, "Origine du fournisseur"] = "fabricant (extract WMS)"
    return d


def annotations() -> pd.DataFrame:
    """Your words, taken as they are, with the sheet they come from."""
    if not DECISIONS.exists():
        return pd.DataFrame(columns=["Code article", "Votre annotation",
                                     "Feuille d'origine"])
    x = pd.ExcelFile(chemin_lisible(DECISIONS))
    morceaux = []
    for feuille in x.sheet_names:
        if feuille.startswith(("0", "1")):
            continue  # sheet 1 decides per supplier, not per item
        d = pd.read_excel(x, sheet_name=feuille, skiprows=3, dtype=str)
        if "Décision" not in d.columns or "Code article" not in d.columns:
            continue
        d = d[["Code article", "Décision"]].copy()
        d["Code article"] = _plein(d["Code article"]).str.replace(
            r"\.0$", "", regex=True)
        d["Votre annotation"] = (_plein(d["Décision"])
                                 .str.replace(r"\s+", " ", regex=True))
        d = d[d["Votre annotation"] != ""]
        d["Feuille d'origine"] = feuille
        morceaux.append(d[["Code article", "Votre annotation",
                           "Feuille d'origine"]])
    if not morceaux:
        return pd.DataFrame(columns=["Code article", "Votre annotation",
                                     "Feuille d'origine"])
    a = pd.concat(morceaux, ignore_index=True)
    return a.drop_duplicates(subset=["Code article"], keep="first")


COLONNES = ["Fournisseur", "Origine du fournisseur", "Code article",
            "Réf. interne", "Désignation", "Famille", "Statut PA",
            "Vendu 12 mois", "Ancien PA", "Nouveau PA",
            "Réf. article fournisseur", "Référence fabricant", "Motif",
            "Votre annotation", "Feuille d'origine"]


def _mettre_en_colonnes(d: pd.DataFrame, motif: pd.Series) -> pd.DataFrame:
    t = d.copy()
    t["Motif"] = motif
    colonnes = [c for c in COLONNES if c in t.columns]
    return t[colonnes]


def classer(d: pd.DataFrame) -> dict:
    """The classification, row by row: THE reference function.

    It works on any table carrying `Famille`, `Fournisseur`,
    `Nom fabriquant` and, if it exists, `Votre annotation`. It reads no
    file: that is what allows `pa_etat_par_article` to call it WHILE it
    is building the base view, at a point where the file `pa_ecartes`
    normally starts from does not exist yet.

    WHY ONE FUNCTION, AND NOT TWO COPIES
        The base view and this module must classify IDENTICALLY,
        otherwise the base view's "Mise à l'écart" column and the sheets
        produced here will contradict each other one day, and nobody will
        know which to believe. A rule written twice is a rule that will
        drift.

    Returns four boolean masks, plus the reason in plain words.
    """
    dit = (_plein(d["Votre annotation"]) if "Votre annotation" in d.columns
           else pd.Series("", index=d.index))
    arret = dit.str.contains(DIT_ARRET) & ~dit.str.contains(DIT_ECARTE)
    ecarte_dit = dit.str.contains(DIT_ECARTE) & ~arret

    # The annotations written in the DELIVERABLE, recorded in the durable
    # register (`pa_annotations_livrable`). They add to those in
    # `pa_decisions.xlsx`: these are two places where the same person
    # writes the same kind of decision, at two different moments.
    codes = _plein(d["Code article"])
    arret |= codes.isin(pa_annotations_livrable.par_lecture("arrêt fabricant"))
    annote_sans_catalogue = codes.isin(
        pa_annotations_livrable.par_lecture("sans catalogue"))

    # ANOTHER SUPPLIER'S NAME INSIDE THE LABEL (rule of 18/09)
    #
    # "STETHOSCOPE MODELE FRSBD GRIS" is attached to Supplier A in the
    # WMS although its label says FRSBD, and its green twin, correctly
    # attached to Supplier BD, does have a price. The item is not
    # untraceable: it is wrongly attached, and we do not fix the WMS from
    # here.
    #
    # We only keep the REAL suppliers of the scope, not just any word: a
    # label containing a common noun that happens to be a brand name
    # would otherwise produce false positives by the series.
    # TWO GUARDRAILS, each paid for by a false positive measured on 18/09:
    #
    #   1. We compare FLATTENED names (no dot, no space). Without that,
    #      "FRSAE <product>" at "FRS.AE MEDICAL" passed for an item of
    #      ANOTHER supplier, because of that one dot: 21 items of
    #      Supplier AE, all already priced, were going into
    #      "discontinued".
    #
    #   2. The rule only holds for items with NO PRICE. A priced item has
    #      no attachment problem to settle: whatever its label says, the
    #      price is there and it is right.
    autre_fournisseur = pd.Series(False, index=d.index)
    if "Désignation" in d.columns:
        libelle = _plein(d["Désignation"]).str.upper()
        aplati = (_plein(d["Fournisseur"]).str.upper()
                  .str.replace(r"[^A-Z0-9]", "", regex=True))
        for marque in MARQUES_FOURNISSEUR:
            porte = libelle.str.contains(rf"\b{marque}\b", regex=True)
            pas_le_sien = ~aplati.str.contains(marque, regex=False)
            autre_fournisseur |= porte & pas_le_sien

    if "Nouveau PA" in d.columns:
        sans_prix = pd.to_numeric(d["Nouveau PA"], errors="coerce").isna()
        autre_fournisseur &= sans_prix

    arret |= autre_fournisseur

    # "IF THEY HAVE A CATALOGUE, THEY HAVE A PRICE": business rule, 18/09
    #
    # This is the rule that closes the loop, and it was already written
    # into the data without being used. `pa_etat_par_article` has always
    # distinguished two reasons for a missing price:
    #
    #   "fournisseur tarifé, mais l'article n'est pas dans son tarif"
    #        -> we DO have the catalogue, the reference is not in it.
    #           So the product has left the range: DISCONTINUED.
    #
    #   "aucun tarif reçu de ce fournisseur"
    #        -> we have NOTHING to read: NO CATALOGUE (handled below).
    #
    # WHAT THIS RESTS ON: the finding is direct, THE SUPPLIER'S CATALOGUE
    # DOES NOT OFFER THIS PRODUCT. We have their price list, we have read
    # it in full, the reference is not there. The program guesses
    # nothing: it matches what it has read and records an absence.
    #
    # The caveat is narrow and concerns isolated cases: a reference that
    # had changed on the supplier's side would produce the same absence.
    # That is why the classification is REDONE at every run and is never
    # frozen: complete the price list or fix the reference in the WMS,
    # and the item comes back into the workstream by itself because its
    # reason will have changed.
    #
    # The reason is only filled in on items without a price (rule of
    # `pa_etat_par_article`): relying on it therefore restricts the reach
    # automatically, with nothing more to say.
    if "Motif" in d.columns:
        motif_lu = _plein(d["Motif"])
        arret |= motif_lu.str.contains("n'est pas dans son tarif",
                                       regex=False)

        # "Aucun fournisseur au périmètre — voir le fabricant": the WMS
        # record does not say who we buy from, only the manufacturer is
        # known. So it is the manufacturer that answers the question "do
        # we have a catalogue?", and the same rule then applies:
        # catalogue + no price = discontinued; no catalogue = no
        # catalogue (further down).
        via_fabricant = motif_lu.str.contains("voir le fabricant",
                                              regex=False)
        if via_fabricant.any():
            fabricant_servi = _plein(d["Nom fabriquant"]).map(a_un_catalogue)
            arret |= via_fabricant & fabricant_servi

    lourd = _plein(d["Famille"]).isin(FAMILLES)
    parque = (_plein(d["Fournisseur"]).str.contains(MIS_DE_COTE)
              | _plein(d["Nom fabriquant"]).str.contains(MIS_DE_COTE))

    # Dropped supplier (Suppliers T, BC...): checked on the supplier AND
    # on the manufacturer. The "4 — HORS CHANTIER" status already says so
    # for most of them, but only where the WMS fills in "Fournisseur"; an
    # item for which only the manufacturer is known (48553) escapes it.
    fournisseur_arrete = (
        _plein(d["Fournisseur"]).map(fournisseurs_arretes.est_arrete)
        | _plein(d["Nom fabriquant"]).map(fournisseurs_arretes.est_arrete))
    if "Statut PA" in d.columns:
        fournisseur_arrete |= _plein(d["Statut PA"]).str.startswith("4")

    # NO CATALOGUE: derived from the reason where it exists, named
    # otherwise.
    #
    # Suppliers Q and Q2 are excluded explicitly: their catalogue exists
    # (`FRSQ-FRSQ2/2026`). If one of their items comes out without a
    # price, then the matching failed, an entirely different problem,
    # which must not be hidden by filing it here. As of 18/09 the
    # registry in fact only looks for "FRSQ": FRSQ2 items are never set
    # against that price list, which is a defect to fix, not a missing
    # source.
    a_un_catalogue_q = (
        _plein(d["Fournisseur"]).str.contains(r"FRSQ|FRSQ2", regex=True)
        | _plein(d["Nom fabriquant"]).str.contains(r"FRSQ|FRSQ2", regex=True))
    sans_catalogue = (_plein(d["Fournisseur"]).str.contains(_SANS_CATALOGUE)
                      | _plein(d["Nom fabriquant"]).str.contains(_SANS_CATALOGUE))
    if "Motif" in d.columns:
        motif_lu = _plein(d["Motif"])
        sans_catalogue |= motif_lu.str.contains("aucun tarif reçu",
                                                regex=False)
        # No supplier in the WMS AND no catalogue at the manufacturer:
        # there is nothing to read anywhere.
        sans_catalogue |= (motif_lu.str.contains("voir le fabricant",
                                                 regex=False)
                           & ~_plein(d["Nom fabriquant"]).map(a_un_catalogue))
    sans_catalogue |= annote_sans_catalogue
    sans_catalogue &= ~a_un_catalogue_q

    # The reason, in order of precedence. The first assignment is the
    # weakest, the last one wins: a wheelchair bought from a supplier
    # with no catalogue is first and foremost a wheelchair, asking for
    # its price list would serve no purpose while it is set aside.
    motif = pd.Series("", index=d.index)
    motif[sans_catalogue] = ("sans catalogue — aucun tarif reçu de "
                             + _plein(d["Fournisseur"])[sans_catalogue]
                             .replace("", "ce fournisseur").str.slice(0, 40))
    motif[parque] = ("fournisseur mis de côté par la direction : "
                     + _plein(d["Fournisseur"])[parque].str.slice(0, 40))
    motif[lourd] = "famille mise de côté : " + _plein(d["Famille"])[lourd]
    motif[ecarte_dit] = "votre annotation"
    motif[arret] = "arrêt fabricant — introuvable au catalogue"
    motif[fournisseur_arrete] = (
        _plein(d["Fournisseur"])[fournisseur_arrete]
        .map(fournisseurs_arretes.motif)
        .where(lambda s: s != "",
              _plein(d["Nom fabriquant"])[fournisseur_arrete]
              .map(fournisseurs_arretes.motif))
        .where(lambda s: s != "", "fournisseur arrêté"))

    return {"arret": arret, "ecarte_dit": ecarte_dit, "lourd": lourd,
            "parque": parque, "fournisseur_arrete": fournisseur_arrete,
            "sans_catalogue": sans_catalogue, "motif": motif}


def construire() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    d = perimetre()
    a = annotations()
    d = d.merge(a, on="Code article", how="left")
    d["Votre annotation"] = _plein(d["Votre annotation"])
    d["Feuille d'origine"] = _plein(d["Feuille d'origine"])

    c = classer(d)
    arret = c["arret"]
    ecarte_dit = c["ecarte_dit"]
    lourd = c["lourd"]
    parque = c["parque"]
    fournisseur_arrete = c["fournisseur_arrete"]
    sans_catalogue = c["sans_catalogue"]
    motif = c["motif"]

    # EACH ITEM IN A SINGLE SHEET, the one for the reason it was given.
    #
    # Without this mutual exclusion, the totals of the four sheets no
    # longer add up: measured on 18/09, they showed 2,225 rows for 1,780
    # items actually set aside, 445 counted twice, because a
    # discontinued wheelchair appeared both in "Écarté" and in "Arrêt
    # fabricant". A reader adding up the tabs then lands on a figure
    # that does not exist.
    #
    # The order is THAT OF THE REASON, so that the sheet an item sits in
    # always matches the label it carries. From the hardest fact to the
    # most revisable choice:
    #     dropped supplier > discontinued > set aside > no catalogue
    arret = arret & ~fournisseur_arrete
    a_ecarter = (ecarte_dit | lourd | parque) & ~(fournisseur_arrete | arret)
    sans_catalogue = sans_catalogue & ~(fournisseur_arrete | arret
                                        | a_ecarter)
    feuille_ecarte = _mettre_en_colonnes(d[a_ecarter], motif[a_ecarter])
    feuille_ecarte = feuille_ecarte.sort_values(
        ["Statut PA", "Fournisseur", "Désignation"],
        ascending=[False, True, True])

    feuille_fournisseur_arrete = _mettre_en_colonnes(
        d[fournisseur_arrete], motif[fournisseur_arrete])
    feuille_fournisseur_arrete = feuille_fournisseur_arrete.sort_values(
        ["Fournisseur", "Désignation"])

    feuille_arret = _mettre_en_colonnes(
        d[arret], pd.Series("introuvable au catalogue fournisseur",
                            index=d.index)[arret])
    feuille_arret = feuille_arret.sort_values(["Fournisseur", "Désignation"])

    feuille_sans_catalogue = _mettre_en_colonnes(
        d[sans_catalogue], motif[sans_catalogue])
    feuille_sans_catalogue = feuille_sans_catalogue.sort_values(
        ["Fournisseur", "Désignation"])

    return (feuille_ecarte, feuille_fournisseur_arrete, feuille_arret,
            feuille_sans_catalogue, d)


def main() -> None:
    ecarte, fournisseur_arrete, arret, sans_catalogue, d = construire()

    chemin = mise_en_forme.chemin_ecriture(SORTIE)
    with pd.ExcelWriter(chemin, engine="openpyxl") as w:
        ecarte.to_excel(w, sheet_name="Écarté", index=False, startrow=3)
        sans_catalogue.to_excel(w, sheet_name="Sans catalogue", index=False,
                                startrow=3)
        fournisseur_arrete.to_excel(w, sheet_name="Fournisseur arrêté",
                                    index=False, startrow=3)
        arret.to_excel(w, sheet_name="Arrêt fabricant", index=False,
                       startrow=3)

    reste = int((ecarte["Statut PA"].str.startswith("3")).sum())
    mise_en_forme.formater(chemin, {
        "Écarté": {
            "ligne_entete": 4, "figer_colonne": 4,
            "titre": "Écarté — lits, fauteuils, releveurs, VPH",
            "sous_titre": (
                f"{len(ecarte)} articles, dont {reste} encore sans prix | "
                f"mis de côté, PAS abandonnés : le produit se vend toujours. "
                f"« Motif » dit pourquoi chacun est là. Arrêté au "
                f"{date.today():%d/%m/%Y}")},
        "Sans catalogue": {
            "ligne_entete": 4, "figer_colonne": 4,
            "titre": "Sans catalogue — aucun tarif reçu du fournisseur",
            "sous_titre": (
                f"{len(sans_catalogue)} articles | ni choix ni arrêt : il "
                f"manque la matière première. Le produit se vend, le "
                f"fournisseur existe, on n'a rien à lire. Le seul geste "
                f"qui les débloque est de réclamer le tarif. Arrêté au "
                f"{date.today():%d/%m/%Y}")},
        "Fournisseur arrêté": {
            "ligne_entete": 4, "figer_colonne": 4,
            "titre": "Fournisseur arrêté — plus de commande chez eux",
            "sous_titre": (
                f"{len(fournisseur_arrete)} articles | constat des appros, "
                f"daté — voir fournisseurs_arretes.py. Le produit n'a rien, "
                f"c'est la relation commerciale qui s'est arrêtée. Arrêté "
                f"au {date.today():%d/%m/%Y}")},
        "Arrêt fabricant": {
            "ligne_entete": 4, "figer_colonne": 4,
            "titre": "Arrêt fabricant — introuvable au catalogue",
            "sous_titre": (
                f"{len(arret)} articles | uniquement ce que VOUS avez "
                f"annoté : le WMS ne signale aucun arrêt (aucun Type 79 au "
                f"périmètre), rien ici n'est déduit. Arrêté au "
                f"{date.today():%d/%m/%Y}")},
    })

    print(f"  {len(ecarte):>5}  Écarté "
          f"({reste} without a price, {len(ecarte) - reste} already priced)")
    for m, n in ecarte["Motif"].str.split(" :").str[0].value_counts().items():
        print(f"         {n:>5}  {m}")
    print(f"  {len(sans_catalogue):>5}  Sans catalogue "
          f"({sans_catalogue['Fournisseur'].nunique()} suppliers)")
    for f, n in sans_catalogue["Fournisseur"].replace(
            "", "(aucun fournisseur)").value_counts().head(10).items():
        print(f"         {n:>5}  {f}")
    print(f"  {len(fournisseur_arrete):>5}  Fournisseur arrêté")
    for f, n in fournisseur_arrete["Fournisseur"].value_counts().items():
        print(f"         {n:>5}  {f}")
    print(f"  {len(arret):>5}  Arrêt fabricant")

    orphelins = d[d["Origine du fournisseur"] == "introuvable"]
    repris = int((d["Origine du fournisseur"] == "fabricant (extract WMS)").sum())
    print(f"\n  {repris} suppliers taken from the manufacturer (WMS extract)")
    if len(orphelins):
        print(f"  {len(orphelins)} with no supplier NOR manufacturer at all:")
        for _, r in orphelins.iterrows():
            print(f"      {r['Code article']:<8} {r['Désignation']}")

    print(f"\n-> {chemin}")


if __name__ == "__main__":
    main()
