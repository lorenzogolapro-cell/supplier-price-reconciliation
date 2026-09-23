# -*- coding: utf-8 -*-
"""
PA workstream: replay a price list match over the frozen scope.

    python pa_completer.py FOURNISSEUR_A
    python pa_completer.py FOURNISSEUR_A FOURNISSEUR_E FOURNISSEUR_I

The price owner integrated 34 price lists and matched each of them
against a population established BEFORE the scope was frozen. On
Supplier A: he processed 576 articles, the scope holds 686, and his
sheet reports 937 "PA hors referencement" - prices found in the price
list and then set aside because the article was not on his list.

This script parses nothing new. It takes the supplier's existing
extractor and matches it against the scope's articles, to recover that
seam of work already paid for.

It follows the price owner's logic, column for column, so that his sheet
can be filled in without any reprocessing:

    Ancien PA | Nouveau PA | Ecart % | Origine Nouveau PA
    Fichier source PA | Alerte

The price owner is authoritative on purchase prices, by governance rule:
this file PROPOSES. It never replaces a value he has already set - the
articles he has processed are flagged and left as they are.

Output: sortie/3-achats/pa_completer_<FOURNISSEUR>.xlsx
    "A completer"   scope articles with no PA, plus the proposed price
    "Deja traite"   what the price owner has already done, for checking
    "Sans tarif"    scope articles absent from the supplier's price list
"""

from __future__ import annotations

import re
import sys
from functools import lru_cache
from pathlib import Path

import pandas as pd

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import eudamed_fiabilite  # noqa: E402
import mise_en_forme  # noqa: E402
import pa_conditionnement  # noqa: E402
import pa_correspondances  # noqa: E402
import pa_identifiants  # noqa: E402
import pa_relecture  # noqa: E402
import pa_remises  # noqa: E402
import perimetre_config  # noqa: E402
import perimetre_liste  # noqa: E402
import prix_achat_wms  # noqa: E402
from achats_par_article import mots_distinctifs  # noqa: E402
from extracteurs import generique  # noqa: E402
from extracteurs.base import (  # noqa: E402
    chemin_lisible,
    cle_ref_souple,
    cle_ref_stricte,
)
from main import (  # noqa: E402
    CATALOGUES,
    FOURNISSEURS,
    SORTIE_ACHATS,
    SORTIE_CHANTIER,
    dernier_prix_achat,
)

INVENTAIRE = SORTIE_CHANTIER / "inventaire_catalogues.xlsx"
from normalisation import normaliser_reference  # noqa: E402

DONNEES = Path("./data")
PERIMETRE = RACINE / "sortie" / "4-perimetre" / "perimetre_v1_1.xlsx"
import wms_extract  # noqa: E402
EXTRACT = wms_extract.chemin()
VERSION = getattr(perimetre_config, "PERIMETRE_VERSION", "v1.1")

# --- alert thresholds, modelled on the price owner's --------------------
# A moderate gap is expected: our prices are old, the price list is this
# year's. What must raise an alert is a gap out of all proportion, which is
# almost always a pack price set against a unit price.
ECART_A_VERIFIER = 0.30       # beyond this, we ask for a review
ECART_SUSPECT = 1.00          # beyond this, the price is probably not comparable
# Ratios that betray a pack size rather than an increase.
CONDITIONNEMENTS = (2, 3, 4, 5, 6, 10, 12, 20, 24, 25, 50, 100)
TOLERANCE_CONDITIONNEMENT = 0.02

# Beyond this ratio, a price increase is no longer a plausible explanation.
# At consumables suppliers, the WMS often stores the price per UNIT
# (EUR 0.0500 per syringe - example value) while the price list quotes per
# CARTON (EUR 5.00 per hundred - example value). The exact ratio does not
# land on a round pack size - carton sizes change, discounts pile on top -
# but that is the cause, and it has to be named rather than settled with a
# "suspect" flag that points nobody anywhere.
RAPPORT_CONDITIONNEMENT_PROBABLE = 3.0

# An IDENTICAL ratio recurring across dozens of different articles is not a
# price increase: it is a discount the price list does not carry. At
# Supplier D, 63 articles come out at 1.2500, which is exactly 1/0.80: the
# price list is gross, our PA is net of a 20% discount (rate and ratio:
# example values).
# A genuine price change never lands on the same ten-thousandth for that
# many references.
# Beyond this ratio to the WMS's PA, the gap can no longer be explained by
# a discount or an increase: it is a price list reading error.
# An x109 at Supplier I came from a pack price taken for a unit price; a
# discount, however steep, never exceeds a factor of 10.
RATIO_ABERRANT = 10.0

REMISE_ARTICLES_MINIMUM = 5      # below this, it is a coincidence
REMISE_ECART_MINIMUM = 0.15      # below this, it is an ordinary increase


def _parfum_conditionnement(rapport) -> int | None:
    """The pack size closest to the ratio, if there is one."""
    if rapport is None or pd.isna(rapport) or rapport <= 0:
        return None
    for n in CONDITIONNEMENTS:
        for cible in (n, 1 / n):
            if abs(rapport - cible) <= TOLERANCE_CONDITIONNEMENT * cible:
                return n
    return None


def alerte(ancien, nouveau) -> str:
    """The alert message, worded the way the price owner words it."""
    if nouveau is None or pd.isna(nouveau):
        return ""
    if ancien is None or pd.isna(ancien) or ancien <= 0:
        return "PAS D'ANCIEN PA — première valorisation"

    rapport = nouveau / ancien
    ecart = rapport - 1
    conditionnement = _parfum_conditionnement(rapport)
    if conditionnement and abs(ecart) > ECART_A_VERIFIER:
        return (f"PA fournisseur NON APPLIQUÉ : {nouveau:.4f} € contre "
                f"{ancien:.4f} € actuel (x{rapport:.1f}) — conditionnement "
                f"probablement différent, à vérifier")
    # Ratio too large to be an increase: we name the cause
    if (rapport >= RAPPORT_CONDITIONNEMENT_PROBABLE
            or rapport <= 1 / RAPPORT_CONDITIONNEMENT_PROBABLE):
        return (f"PA fournisseur NON APPLIQUÉ : {nouveau:.4f} € contre "
                f"{ancien:.4f} € actuel (x{rapport:.1f}) — écart trop fort "
                f"pour une hausse ; prix au colis contre prix unitaire ?")
    if abs(ecart) > ECART_SUSPECT:
        return f"VÉRIF ÉCART {ecart:+.0%} — SUSPECT"
    if abs(ecart) > ECART_A_VERIFIER:
        return f"VÉRIF ÉCART {ecart:+.0%} — INCOHÉRENCE POSSIBLE"
    return ""


# ---------------------------------------------------------------------------

# The WMS name and the catalogue folder sometimes have NO word in common:
# no mechanism can bring them together, so it has to be stated once.
ALIAS_CATALOGUE = {
    # On the left, the company name as the WMS carries it; on the right,
    # the name of the catalogue folder: the two share no word at all.
    "FOURNISSEUR AM": "DOSSIER-AM",
    "FOURNISSEUR BX": "DOSSIER-BX",
    # For Supplier BK, the company name is the legal entity's while the
    # folder carries its commercial BRAND: the WMS knows the first, the
    # catalogue folder carries the second. Without this alias, that
    # supplier's 27 articles pass for "no price list received" even though
    # the 2026 price list is right there.
    "FOURNISSEUR BK": "DOSSIER-BK",
}

# Words too widespread in this industry's company names to identify
# anybody. Without this filter, Supplier BV's company name - which contains
# CONFORT - lands on Supplier CT's price list, which contains it too: the
# very mistake the price owner's tracking sheet made as well.
MOTS_TROP_COURANTS = {
    "CONFORT", "SANTE", "MEDICAL", "MEDICALE", "FRANCE", "FRANCAISE",
    "SAS", "SARL", "GROUPE", "LABORATOIRE", "LABORATOIRES", "MEDICAUX",
    "INTERNATIONAL", "INTERNATIONALE", "DISTRIBUTION", "TECHNIQUE",
    "SERVICES", "EUROPE", "PHARMA", "SELF", "TOUS",
    # Supplier BV's company name landed on Supplier AX's folder through the
    # single word ERGO - two different companies. ERGO describes ergonomics,
    # not a trading name.
    "ERGO", "CONCEPT",
}


def mots_du_fournisseur(nom) -> set:
    """Identifying words of a company name, filler words removed."""
    return mots_distinctifs(nom) - MOTS_TROP_COURANTS


def tarif_du_fournisseur(motif_wms: str,
                         noms_reels: tuple[str, ...] = ()) -> pd.DataFrame:
    """Price list rows for the supplier, via the existing extractor.

    We rewrite no parser. Two families of suppliers:

      - those in main.py's FOURNISSEURS registry, which have a dedicated
        extractor that knows their pricing logic;
      - all the others, read by extracteurs/generique from the inventory -
        that is how consolider.py picks up Suppliers E, CS or CD, which are
        absent from the registry.
    """
    # The words that identify this supplier: those of its real WMS name,
    # plus those of its catalogue folder when the two share no word (the
    # folder "DOSSIER-AM" against Supplier AM's company name).
    attendus = set()
    for nom in noms_reels or (motif_wms,):
        attendus |= mots_du_fournisseur(nom)
    for alias, dossier in ALIAS_CATALOGUE.items():
        if any(alias in str(n).upper() for n in (noms_reels or (motif_wms,))):
            attendus |= mots_du_fournisseur(dossier)
    if not attendus:
        return pd.DataFrame()

    lots = []
    for entree in FOURNISSEURS:
        # The registry is tested word by word, not by equality: its
        # motif_wms is an in-house fragment that does not always match the
        # WMS label.
        if not (mots_du_fournisseur(entree["motif_wms"]) & attendus
                or mots_du_fournisseur(entree["dossier"]) & attendus):
            continue
        dossier = CATALOGUES / entree["dossier"]
        for fichier in sorted(dossier.glob(entree["motif"])):
            if fichier.name.startswith("~$"):
                continue
            try:
                lot = entree["extracteur"].extract(fichier)
            except Exception as err:
                print(f"    ! {fichier.name} : {err}")
                continue
            lot["fichier_tarif"] = fichier.name
            lots.append(lot)
            print(f"    dedicated - {fichier.name} : {len(lot)} rows")
    if lots:
        return pd.concat(lots, ignore_index=True)

    # Generic fallback, over the files the inventory kept.
    #
    # Matching is done on DISTINCTIVE WORDS, never on substrings: "PLAST"
    # occurs inside Supplier O's company name, and an early attempt did
    # apply its price list to Supplier EA, whose company name contains
    # "PLASTIQUE". A substring says nothing about identity.
    if not INVENTAIRE.exists():
        print("    ! inventory missing: run inventaire_catalogues.py")
        return pd.DataFrame()
    inventaire = pd.read_excel(chemin_lisible(INVENTAIRE),
                               sheet_name="Retenus", dtype=str)
    excels = inventaire[inventaire["format"].isin(("xlsx", "xlsm", "xls"))]
    for _, ligne in excels.iterrows():
        if not mots_du_fournisseur(ligne["fournisseur"]) & attendus:
            continue
        chemin = Path(ligne["chemin"])
        if not chemin.exists():
            continue
        try:
            lot = generique.extract(chemin, fournisseur=ligne["fournisseur"])
        except Exception as err:
            print(f"    ! {chemin.name} : {type(err).__name__}")
            continue
        if lot.empty:
            continue
        lot["fichier_tarif"] = chemin.name
        lots.append(lot)
        print(f"    generic - {chemin.name} : {len(lot)} rows")
    if not lots:
        return pd.DataFrame()
    return pd.concat(lots, ignore_index=True)


COMMANDES = DONNEES / "lignes_de_commande.xlsx"


@lru_cache(maxsize=1)
def depuis_les_commandes() -> pd.DataFrame:
    """Supplier and supplier reference, taken from past purchase orders.

    387 articles in the scope have NO supplier on their article record,
    while 277 of them have one - and only one - on their order lines. The
    supplier is not unknown, it is elsewhere: without this fallback, those
    articles attach to no price list and drop out of any reasoning by
    supplier.

    We keep the MOST RECENT order: that is the one that says who we buy
    from today, since a supplier may have changed over two years.
    """
    if not COMMANDES.exists():
        print("    ! purchase order history missing")
        return pd.DataFrame()
    cmd = pd.read_excel(chemin_lisible(COMMANDES), dtype=str,
                        usecols=["Réf. Art.", "Fournisseur", "Réf. Four.",
                                 "Date de Cde"])
    cmd = cmd.dropna(subset=["Réf. Art.", "Fournisseur"])
    cmd["cle_interne"] = cmd["Réf. Art."].map(normaliser_reference)
    cmd["date"] = pd.to_datetime(cmd["Date de Cde"], errors="coerce")
    # "229 - FOURNISSEUR Y" -> "FOURNISSEUR Y"
    cmd["fournisseur_cmd"] = (cmd["Fournisseur"].str.split("-", n=1)
                              .str[-1].str.strip())
    cmd = cmd.sort_values("date").drop_duplicates("cle_interne", keep="last")
    return cmd[["cle_interne", "fournisseur_cmd", "Réf. Four."]].rename(
        columns={"Réf. Four.": "ref_fournisseur_cmd"})


def references_en_collision(dedans: pd.DataFrame) -> set:
    """Internal references carried by several articles in the scope.

    The Reference is not unique: 7 of them carry two articles, and one does
    so across TWO different suppliers - a water bottle at Supplier A and an
    electrode clip at Supplier AP. Any match on that key would give them
    the same thing, and would be wrong half the time.

    We do not decide: we list them so as to keep them out of the fallback
    and report them in a column.
    """
    utiles = dedans[dedans["cle_interne"].notna()
                    & (dedans["cle_interne"] != "")]
    comptes = utiles["cle_interne"].value_counts()
    return set(comptes[comptes > 1].index)


@lru_cache(maxsize=1)
def referentiel_enrichi() -> pd.DataFrame:
    """The frozen scope, completed with its supplier reference.

    Cached: this load reads three workbooks of several thousand rows each,
    and the script is called supplier by supplier - without the cache it
    would reread them on every pass.
    """
    # The scope comes from the frozen list, never from a recomputation: two
    # sessions that recompute can diverge without anyone noticing.
    dedans = perimetre_liste.charger().copy()
    print(f"    {perimetre_liste.entete()}")
    dedans = dedans.rename(columns={"Fournisseur": "Nom fournisseur",
                                    "Désignation": "Libellé déclinaison ^(1)"})
    dedans["cle_interne"] = dedans["Réf. interne"].map(normaliser_reference)

    # The Reference is not unique: we flag the cases rather than let them
    # silently skew a match.
    collisions = references_en_collision(dedans)
    dedans["collision_reference"] = dedans["cle_interne"].isin(collisions)
    sans_reference = (dedans["cle_interne"].isna()
                      | (dedans["cle_interne"] == ""))
    dedans["sans_reference"] = sans_reference
    if collisions or sans_reference.any():
        print(f"    Reference key: {len(collisions)} references carried by "
              f"several articles, {int(sans_reference.sum())} articles with "
              f"no reference - excluded from the fallback, flagged in a column")

    # The supplier reference comes from the article extract: it is what
    # bridges to the price list. The WMS ALSO carries a manufacturer pair
    # (name and reference), far more complete: 384 of the 387 articles
    # without a supplier have a manufacturer, and 1,846 of the articles
    # without a price have a manufacturer reference. That is a second door
    # into the price list.
    art = pd.read_excel(chemin_lisible(EXTRACT), dtype=str,
                        usecols=["Code article", "Ref. art. four.",
                                 "Nom fabriquant", "Référence fabricant"])
    pont = art.dropna(subset=["Code article"]).drop_duplicates("Code article")
    codes = pont["Code article"].str.strip()
    cle = dedans["Code article"].str.strip()
    dedans["ref_fournisseur_wms"] = cle.map(
        dict(zip(codes, pont["Ref. art. four."])))
    dedans["ref_fabricant"] = cle.map(
        dict(zip(codes, pont["Référence fabricant"])))
    dedans["nom_fabriquant"] = cle.map(
        dict(zip(codes, pont["Nom fabriquant"])))

    # Fallback on the order history, for incomplete article records
    cmd = depuis_les_commandes()
    if not cmd.empty:
        dedans = dedans.merge(cmd, on="cle_interne", how="left")
        vide = dedans["Nom fournisseur"].fillna("").str.strip() == ""
        dedans["origine_fournisseur"] = "fiche article"
        # A colliding reference designates two articles: the fallback would
        # give them the same supplier, wrong for one of the two.
        exploitable = ~dedans["collision_reference"] & ~dedans["sans_reference"]
        recuperes = vide & exploitable & dedans["fournisseur_cmd"].notna()
        dedans.loc[recuperes, "Nom fournisseur"] = \
            dedans.loc[recuperes, "fournisseur_cmd"]
        dedans.loc[recuperes, "origine_fournisseur"] = "historique commande"
        # Same for the reference, often missing as well
        sans_ref = dedans["ref_fournisseur_wms"].fillna("").str.strip() == ""
        repris = sans_ref & exploitable & dedans["ref_fournisseur_cmd"].notna()
        dedans.loc[repris, "ref_fournisseur_wms"] = \
            dedans.loc[repris, "ref_fournisseur_cmd"]
        if recuperes.any() or repris.any():
            print(f"    order fallback: {int(recuperes.sum())} suppliers "
                  f"and {int(repris.sum())} references recovered")
    else:
        dedans["origine_fournisseur"] = "fiche article"

    # Last fallback: the manufacturer. It does NOT say who we buy from - we
    # may buy Supplier B's goods through a distributor - but without it
    # these articles attach to no price list at all. The origin stays in a
    # column so that the nuance is visible at review time.
    vide = dedans["Nom fournisseur"].fillna("").str.strip() == ""
    depuis_fabricant = vide & dedans["nom_fabriquant"].fillna(
        "").str.strip().ne("")
    dedans.loc[depuis_fabricant, "Nom fournisseur"] = \
        dedans.loc[depuis_fabricant, "nom_fabriquant"]
    dedans.loc[depuis_fabricant, "origine_fournisseur"] = \
        "fabricant (fiche article)"
    if depuis_fabricant.any():
        print(f"    manufacturer fallback: {int(depuis_fabricant.sum())} "
              f"suppliers recovered - attribution to be confirmed")
    return dedans


def articles_du_perimetre(motif_wms: str) -> pd.DataFrame:
    """Frozen scope articles belonging to this supplier."""
    dedans = referentiel_enrichi()
    sien = dedans[dedans["Nom fournisseur"].fillna("").str.upper()
                  .str.contains(motif_wms.upper(), regex=False)].copy()
    sien["cle_stricte"] = sien["ref_fournisseur_wms"].map(cle_ref_stricte)
    sien["cle_souple"] = sien["ref_fournisseur_wms"].map(cle_ref_souple)
    sien["cle_fabricant"] = sien["ref_fabricant"].map(cle_ref_stricte)
    return sien


@lru_cache(maxsize=1)
def ean_par_article() -> pd.DataFrame:
    """Article code -> its reliable EAN codes, from the EAN workstream.

    Fourth door into the price list, and the most universal one: an EAN
    depends neither on the supplier's numbering nor on ours. Many price
    lists publish it even when their reference has changed.

    We only keep EANs declared RELIABLE. A "probable" EAN that lands on the
    wrong price list line does not produce an error: it produces a price, a
    wrong one, and nobody will see it.

    Two fixes made on 15/09, on this same function:

    The file. `sorted(glob("EAN distributeur*.xlsx"))[-1]` was picking
    "EAN distributeur.xlsx", the WORKING file, the one that changes with
    every collection run and that the convention forbids citing.
    Alphabetical order puts "." after "-", and in any case it would have
    put "V9" after "V16". We sort on the number instead.

    The filter. "Fiabilité == sûr" was letting through 1,321 EUDAMED codes
    obtained on a reference too short to designate anything at all.
    Checked against the label in EUDAMED itself, 3 out of 370 held up. They
    only matched 33 prices, 5 of them through a dubious link - the 4th key
    is of little use - but a wrong, silent price is exactly what this
    function says it wants to avoid.
    """
    dossier = RACINE / "sortie" / "2-chantier-ean"
    fichiers = sorted(
        dossier.glob("EAN distributeur - V*.xlsx"),
        key=lambda p: int(re.search(r"- V(\d+)", p.name).group(1)),
    )
    if not fichiers:
        return pd.DataFrame(columns=["Code article", "cle_ean"])
    table = pd.read_excel(chemin_lisible(fichiers[-1]),
                          sheet_name="Codes EAN", skiprows=3, dtype=str)
    table = table[table["Fiabilité"].fillna("").str.strip() == "sûr"]
    table = table[~eudamed_fiabilite.douteux(table)]
    table = table.dropna(subset=["Code article", "Code EAN"])
    liens = pd.DataFrame({
        "Code article": table["Code article"].str.strip(),
        "cle_ean": table["Code EAN"].str.strip(),
    }).drop_duplicates()
    print(f"    EAN reference data: {len(liens)} reliable links "
          f"over {liens['Code article'].nunique()} articles")
    return liens


@lru_cache(maxsize=1)
def pa_actuels() -> pd.DataFrame:
    """The WMS's last purchase price, by resolved article code."""
    chemin = dernier_prix_achat()
    if chemin is None:
        return pd.DataFrame(columns=["code_article", "pa_wms"])
    pa = prix_achat_wms.charger(chemin)
    pa = pa.dropna(subset=["code_article"])
    return pa.drop_duplicates("code_article")[["code_article", "pa_wms"]]


@lru_cache(maxsize=1)
def deja_traites_par_le_responsable() -> frozenset:
    """Internal keys the price owner has already given a Nouveau PA."""
    chemin = DONNEES / "pa_valides.xlsx"
    if not chemin.exists():
        return frozenset()
    valides = pd.read_excel(chemin_lisible(chemin), sheet_name="Feuil1",
                            dtype=str).dropna(subset=["Code article"])
    valides["pa"] = pd.to_numeric(valides["Nouveau PA"], errors="coerce")
    valides = valides[valides["pa"].fillna(0) > 0]

    art = pd.read_excel(chemin_lisible(EXTRACT), dtype=str,
                        usecols=["Code article", "Référence"])
    pont = art.dropna(subset=["Code article"]).drop_duplicates("Code article")
    index = dict(zip(pont["Code article"].str.strip(),
                     pont["Référence"].map(normaliser_reference)))
    return frozenset(valides["Code article"].str.strip().map(index).dropna())


COLONNES_SORTIE = [
    "PERIMETRE_VERSION", "Réf. interne", "Code article",
    "Libellé déclinaison ^(1)", "Nom fournisseur", "origine_fournisseur",
    "Type", "CATEGORIE", "Réf. fournisseur", "Désignation tarif",
    "Ancien PA", "Nouveau PA", "Écart %",
    "Origine Nouveau PA", "Fichier source PA", "Alerte",
    "PA écarté (à vérifier)", "Paliers", "Conditionnement",
    # A price brought back to the unit must carry the trace of how that was
    # done: by which divisor, on what proof, and what the price was before.
    # Without these columns, a wrong division cannot be detected afterwards.
    "diviseur_conditionnement", "source_diviseur", "preuve_diviseur",
    "prix_avant_division", "config_conditionnement", "minimum_de_commande",
    # And the trace of what we REFUSED to divide: a pack size seen in a
    # label and then set aside because the DPA did not support it. Without
    # this column, the row looks like a row no rule ever touched, and
    # nobody knows a judgement call was made.
    "division_ecartee",
    # Same requirement for a discount absent from the price list: the rate
    # applied, the price before discount, and the reason when it was set
    # aside. And when a FAMILY escapes the supplier's general rate - MODELE
    # EXEMPLE 1 at 25% where Supplier D is at 20%, example values - the
    # column says so with
    # its measurement. A row that does not follow the general rule must be
    # visible.
    "remise_taux_applique", "prix_avant_remise", "remise_ecartee",
    "remise_famille",
    # The key is not always reliable: we say so in a column rather than let
    # people believe in a clean match.
    "Clé de rapprochement", "preuve_identifiant",
    "collision_reference", "sans_reference",
]


def completer(motif_wms: str,
              articles: pd.DataFrame | None = None) -> dict[str, pd.DataFrame]:
    """Matches a supplier's articles against its price list.

    `articles` makes it possible to work on a DIFFERENT population from the
    frozen scope - the professional catalogue, for instance, part of which
    had neither a sale nor a receipt over twelve months and therefore does
    not enter the scope. The scope stays what it is: we do not widen it, we
    set a population alongside it, carrying its own label.
    """
    print(f"\n{'=' * 74}\n{motif_wms}\n{'=' * 74}")

    # We identify the articles first, in order to know WHICH supplier this
    # really is. Looking for the price list from the fragment passed as an
    # argument alone means risking a namesake.
    if articles is None:
        articles = articles_du_perimetre(motif_wms)
    print(f"  population: {len(articles)} articles")
    if articles.empty:
        print("  no article in the scope - nothing to do here")
        return {}
    noms = tuple(sorted(articles["Nom fournisseur"].dropna().unique()))
    print(f"  suppliers targeted: {', '.join(n[:34] for n in noms[:4])}")
    if len(noms) > 1:
        print(f"  ! the fragment \"{motif_wms}\" designates {len(noms)} "
              f"distinct suppliers - pass a more precise fragment")

    print("  price list:")
    tarif = tarif_du_fournisseur(motif_wms, noms)
    if tarif.empty:
        print("  no usable price list - nothing to do here")
        return {}

    # Strict then loose matching, as everywhere in the project
    champs = ["ref_fournisseur", "designation", "prix_achat_unitaire_ht",
              "paliers", "conditionnement", "fichier_tarif", "colonne_prix"]
    champs = [c for c in champs if c in tarif.columns]
    tarif = tarif.copy()
    tarif["cle_stricte"] = tarif["ref_fournisseur"].map(cle_ref_stricte)
    tarif["cle_souple"] = tarif["ref_fournisseur"].map(cle_ref_souple)

    strict = tarif.dropna(subset=["cle_stricte"]).drop_duplicates("cle_stricte")
    fusion = articles.merge(strict[["cle_stricte"] + champs],
                            on="cle_stricte", how="left",
                            suffixes=("", "_tarif"))
    fusion["Clé de rapprochement"] = ""
    fusion.loc[fusion["prix_achat_unitaire_ht"].notna(),
               "Clé de rapprochement"] = "réf. fournisseur"

    # Two fallback passes. The loosened supplier reference first, then the
    # MANUFACTURER reference: on the articles still without a price, it is
    # filled in exactly where the supplier reference is missing. A weaker
    # key than the previous one, so it is named in a column.
    for cle_article, cle_tarif, etiquette in (
            ("cle_souple", "cle_souple", "réf. fournisseur (souple)"),
            ("cle_fabricant", "cle_stricte", "réf. fabricant")):
        absents = fusion["prix_achat_unitaire_ht"].isna()
        if not absents.any() or cle_article not in articles.columns:
            continue
        source = (tarif.dropna(subset=[cle_tarif])
                  .drop_duplicates(cle_tarif)
                  .rename(columns={cle_tarif: cle_article}))
        appoint = articles[absents.values].merge(
            source[[cle_article] + champs], on=cle_article, how="left",
            suffixes=("", "_tarif"))
        for champ in champs:
            fusion.loc[absents.values, champ] = appoint[champ].values
        neufs = absents & fusion["prix_achat_unitaire_ht"].notna()
        fusion.loc[neufs, "Clé de rapprochement"] = etiquette
        if neufs.any():
            print(f"  {int(neufs.sum())} matched by {etiquette}")

    # Fourth pass: the EAN. It cannot work like the other three - an
    # article often carries SEVERAL EAN codes, so the join duplicates rows.
    # We match separately, keep only one price list row per article, then
    # carry it back by article code.
    absents = fusion["prix_achat_unitaire_ht"].isna()
    liens = ean_par_article()
    if absents.any() and "ean" in tarif.columns and not liens.empty:
        source = tarif.copy()
        source["cle_ean"] = source["ean"].astype(str).str.strip()
        source = source[source["cle_ean"].str.len() >= 8]
        source = source.drop_duplicates("cle_ean")
        if not source.empty:
            candidats = (fusion.loc[absents, ["Code article"]]
                         .merge(liens, on="Code article")
                         .merge(source[["cle_ean"] + champs], on="cle_ean")
                         .drop_duplicates("Code article")
                         .set_index("Code article"))
            vise = absents & fusion["Code article"].isin(candidats.index)
            for champ in champs:
                fusion.loc[vise, champ] = (
                    fusion.loc[vise, "Code article"].map(candidats[champ]))
            fusion.loc[vise, "Clé de rapprochement"] = "code EAN"
            if vise.any():
                print(f"  {int(vise.sum())} matched by EAN code")

    # Fifth pass: damaged identifiers. It comes AFTER the four keys and
    # BEFORE any fuzzy matching, because an EAN whose check digit falls
    # into place is an EXACT match, not a candidate.
    fusion = pa_identifiants.rapprocher_par_identifiant(fusion, tarif, champs)

    # Sixth pass: the matches a human established by hand, one by one, each
    # with its proof.
    fusion = pa_correspondances.appliquer(fusion, tarif, champs, motif_wms)

    # Seventh and last pass: the "OUI" decisions from pa_a_relire.xlsx, a
    # batch of label candidates validated together on 17/09, after
    # filtering out orphan lots/sets and hardware conflations. Same rule:
    # never overwrite a price already set.
    fusion = pa_relecture.appliquer(fusion, motif_wms)

    pa = pa_actuels()
    fusion = fusion.merge(pa, left_on=fusion["Code article"].str.strip(),
                          right_on="code_article", how="left")

    # Pack size: some suppliers quote by the lot while the WMS counts by
    # the unit. The DPA is joined JUST BEFORE, because inference from
    # systematic recurrence needs it - it is the ratio to the DPA that
    # reveals a repeated factor. Anything divided says so in a column.
    fusion = pa_conditionnement.appliquer_au_rapprochement(fusion)

    # Discount absent from the price list: some suppliers publish a GROSS
    # price and the negotiated discount appears nowhere in the file. AFTER
    # the pack size, never before: we discount a price already brought back
    # to the unit. The rate comes not from a document but from the prices
    # the price owner arbitrated, and it only applies if the DPA
    # corroborates it.
    fusion = pa_remises.appliquer_au_rapprochement(fusion)

    # Four decimals, like the price owner's workbook: otherwise a computed
    # price drags its floating-point decimals along (12.3399999999999 -
    # example value), which are
    # not a price but an artefact.
    fusion["Ancien PA"] = pd.to_numeric(fusion["pa_wms"],
                                        errors="coerce").round(4)
    fusion["Nouveau PA"] = pd.to_numeric(fusion["prix_achat_unitaire_ht"],
                                         errors="coerce").round(4)
    fusion["Écart %"] = (
        (fusion["Nouveau PA"] - fusion["Ancien PA"]) / fusion["Ancien PA"]
    ).round(4)
    fusion["Origine Nouveau PA"] = (
        "Tarif fournisseur " + fusion["fichier_tarif"].fillna("")
    ).where(fusion["Nouveau PA"].notna(), "")
    fusion["Fichier source PA"] = fusion["fichier_tarif"].fillna("")
    fusion["Alerte"] = [alerte(a, n) for a, n
                        in zip(fusion["Ancien PA"], fusion["Nouveau PA"])]

    # Systematic discount: the same ratio across many articles says the
    # price list is GROSS and our PA net. It is spotted by recurrence, not
    # by magnitude - that is what distinguishes it from an increase.
    ancien = pd.to_numeric(fusion["Ancien PA"], errors="coerce")
    nouveau = pd.to_numeric(fusion["Nouveau PA"], errors="coerce")
    ratio = (nouveau / ancien).round(2)
    comptes = ratio[ratio.notna() & ((ratio - 1).abs() > REMISE_ECART_MINIMUM)]
    suspects = {r for r, n in comptes.value_counts().items()
                if n >= REMISE_ARTICLES_MINIMUM}
    remise = ratio.isin(suspects)
    if remise.any():
        for r in sorted(suspects):
            n = int((ratio == r).sum())
            fusion.loc[ratio == r, "Alerte"] = (
                f"écart systématique ×{r:.2f} sur {n} articles — le tarif "
                f"est plus cher que le WMS de {1 - 1 / r:.0%} : hausse, ou "
                f"remise que le tarif ne porte pas ?" if r > 1 else
                f"écart systématique ×{r:.2f} sur {n} articles — tarif et "
                f"PA ne semblent pas sur la même base")

    # The price list price is kept even when the gap is systematic: the
    # column read really is the unit price or the discounted price, never
    # the list price - that is the purchasing rule. A recurring gap
    # therefore becomes an alert to review, not grounds for refusal.
    #
    # One guard rail remains, because not every gap is a discount: a ratio
    # of x109 at Supplier I came from a pack price taken for a unit price.
    # Beyond RATIO_ABERRANT it is no longer a commercial question but a
    # reading error, and that one does not get injected.
    aberrant = ratio.notna() & ((ratio > RATIO_ABERRANT)
                                | (ratio < 1 / RATIO_ABERRANT))
    fusion.loc[aberrant, "Alerte"] = (
        "rapport ×" + ratio[aberrant].round(1).astype(str)
        + " avec le PA du WMS — trop grand pour une remise, "
          "prix au colis lu comme un prix unitaire ?")
    fusion["PA écarté (à vérifier)"] = fusion["Nouveau PA"].where(aberrant)
    fusion.loc[aberrant, "Nouveau PA"] = None
    fusion.loc[aberrant, "Écart %"] = None
    ecartes = aberrant
    if remise.any():
        print(f"  {int(remise.sum())} values with a systematic gap, "
              f"kept with an alert")
    if aberrant.any():
        print(f"  {int(aberrant.sum())} values set aside "
              f"(aberrant ratio, > x{RATIO_ABERRANT:.0f})")

    fusion["PERIMETRE_VERSION"] = VERSION
    fusion = fusion.rename(columns={
        "ref_fournisseur": "Réf. fournisseur",
        "designation": "Désignation tarif",
        "paliers": "Paliers",
        "conditionnement": "Conditionnement",
    })

    servies = deja_traites_par_le_responsable()
    fusion["deja_traite"] = fusion["cle_interne"].isin(servies)
    trouve = fusion["Nouveau PA"].notna()
    a_faire = ~fusion["deja_traite"]

    # A price set aside is NOT a price absent from the price list: it is in
    # there, we simply refuse to apply it until the pack size is checked.
    # Mixing the two would make both sheets wrong.
    lots = {
        "À compléter": fusion[a_faire & trouve],
        "Déjà traité": fusion[fusion["deja_traite"]],
        "Écartés": fusion[a_faire & ~trouve & ecartes],
        "Sans tarif": fusion[a_faire & ~trouve & ~ecartes],
    }
    # Governance rule: the total must be preserved, every row accounted for
    total = sum(len(t) for t in lots.values())
    assert total == len(fusion), f"total not preserved: {total} / {len(fusion)}"

    print(f"  to complete: {len(lots['À compléter'])}")
    print(f"  already handled by the price owner: "
          f"{len(lots['Déjà traité'])}")
    print(f"  absent from the price list: {len(lots['Sans tarif'])}")
    alertes = lots["À compléter"]["Alerte"]
    if len(alertes):
        posees = (alertes != "").sum()
        print(f"  of which {posees} carry an alert")

    for nom in lots:
        lots[nom] = lots[nom].reindex(
            columns=[c for c in COLONNES_SORTIE if c in fusion.columns])
    return lots


def _nom_fichier(nom: str) -> str:
    """A safe file name, and above all a DISTINCT one across suppliers.

    Taking the first three words gave the same name to two different
    entities: "FOURNISSEUR BF FRANCE SITE DE ALPHA" and "FOURNISSEUR BF
    FRANCE SITE DE BETA" both produced FOURNISSEUR-BF-FRANCE. The second
    overwrote the first, and the sweep then believed it had processed it -
    36 articles lost without a sound.

    So we keep the IDENTIFYING words, the ones left once the industry's
    widespread words are removed: ALPHA and BETA survive, FRANCE and SITE
    disappear.
    """
    mots = [m for m in re.split(r"[^A-Za-z0-9]+", nom.upper()) if m]
    distinctifs = [m for m in mots if m not in MOTS_TROP_COURANTS
                   and len(m) > 2]
    retenus = distinctifs[:4] or mots[:3]
    return "-".join(retenus)[:60] or "FOURNISSEUR"


def fournisseurs_restants() -> list[str]:
    """Scope suppliers that no pass has covered yet.

    A full sweep beats a hand-picked list: of the seven suppliers we
    thought only had a PDF price list, three in fact had an Excel one
    already read by the project. Asking the extractor costs nothing;
    guessing costs articles.
    """
    # The names come from the ENRICHED reference data, not from the raw
    # scope: otherwise the articles attached through the order history or
    # through the manufacturer designate a supplier the sweep never opens.
    dedans = referentiel_enrichi()
    noms = (dedans["Nom fournisseur"].dropna().astype(str).str.strip())
    noms = noms[noms != ""]

    # Comparison on the EXACT file name, not on common words. The
    # intersection test treated "FOURNISSEUR BF ... ALPHA" as already
    # processed because a file existed for "FOURNISSEUR BF ... BETA": two
    # distinct entities, a single sweep.
    deja = {f.stem.replace("pa_completer_", "").upper()
            for f in SORTIE_ACHATS.glob("pa_completer_*.xlsx")}
    restants = []
    for nom, n in noms.value_counts().items():
        if _nom_fichier(nom).upper() in deja:
            continue
        restants.append(nom)
    print(f"{len(restants)} suppliers left to sweep\n")
    return restants


def main() -> None:
    arguments = sys.argv[1:]
    if arguments and arguments[0] == "--tous":
        cibles = fournisseurs_restants()
    else:
        cibles = arguments or ["FOURNISSEUR_A"]
    for motif in cibles:
        lots = completer(motif)
        if not lots:
            continue
        sortie = mise_en_forme.chemin_ecriture(
            SORTIE_ACHATS / f"pa_completer_{_nom_fichier(motif)}.xlsx")
        with pd.ExcelWriter(sortie, engine="openpyxl") as writeur:
            for nom, table in lots.items():
                table.to_excel(writeur, sheet_name=nom, index=False,
                               startrow=3)
        mise_en_forme.formater(sortie, {
            "À compléter": {
                "ligne_entete": 4, "figer_colonne": 4,
                "titre": f"{motif} — PA à compléter sur le périmètre {VERSION}",
                "sous_titre": "Proposition. Le responsable des prix fait "
                              "foi : on comble, on ne remplace pas."},
            "Déjà traité": {
                "ligne_entete": 4, "figer_colonne": 4,
                "titre": "Déjà doté d'un Nouveau PA par le responsable des prix",
                "sous_titre": "Pour contrôle — ne rien y changer"},
            "Écartés": {
                "ligne_entete": 4, "figer_colonne": 4,
                "titre": "Prix trouvé au tarif mais NON APPLIQUÉ",
                "sous_titre": "Écart trop fort pour une hausse — prix au "
                              "colis contre prix unitaire ? Valeur conservée "
                              "en colonne « PA écarté »"},
            "Sans tarif": {
                "ligne_entete": 4, "figer_colonne": 4,
                "titre": "Au périmètre mais absents du tarif fournisseur",
                "sous_titre": "Ni un bug ni un oubli : le tarif ne les porte pas"},
        })
        print(f"  -> {sortie}")


if __name__ == "__main__":
    main()

