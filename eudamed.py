# -*- coding: utf-8 -*-
"""
Queries EUDAMED, the European medical device database.

Every manufacturer placing a medical device on the European market has
to declare it there with its UDI-DI. And a UDI-DI IS a GTIN: this is
therefore a public source of EAN codes, including for suppliers that
publish none in their price lists.

    EAN-13 = primaryDi (GTIN-14) stripped of its leading zero

The API is public and needs no authentication, but it is temperamental:

  - the "srn" filter (manufacturer registration number) works and
    returns that manufacturer's entire catalogue;
  - the "tradeName" filter works too, as a partial, case-insensitive
    search;
  - ALL THE OTHERS are silently ignored: "manufacturerName", "keyword",
    "query"... The API then answers 200 with the 3.3 million devices of
    the whole database. A total above 100 000 therefore signals a filter
    that was not taken into account, not a result.

There is no public endpoint to search for actors: a manufacturer's SRN
is discovered either through one of its products (via tradeName), or in
its CE certificates.
"""

from __future__ import annotations

import json
import re
import time
import unicodedata
from pathlib import Path

import requests

URL = "https://ec.europa.eu/tools/eudamed/api/devices/udiDiData"
ENTETES = {
    "Accept": "application/json",
    "User-Agent": "catalogue-distributeur/1.0 (mise a jour base articles)",
}

# Above this, the requested filter was ignored and the API returns the
# whole database
SEUIL_FILTRE_IGNORE = 100_000

# Politeness towards a public service: a pause between two calls
PAUSE = 0.4
TAILLE_PAGE = 100

# Known SRNs, to avoid a discovery pass on every run.
# The SRN can also be read in the CE certificates published by the
# manufacturer.
# SRNs verified one by one: we query EUDAMED by the trade name of a known
# product, then read the company name of the manufacturer that answers.
# Searching by company name, on the other hand, returns nothing: that is
# a limit of the API, not a missing declaration.
#
# Suppliers P, L and AM are not listed here: the searches only bring back
# homonyms (suppliers GC, CJ and GD). No point retrying.
SRN_CONNUS = {
    "FOURNISSEUR B": "BE-MF-000000001",
    "FOURNISSEUR H": "FR-MF-000000002",   # supplier H's company name
    "FOURNISSEUR Q": "FR-MF-000000003",   # supplier Q's company name
    "FOURNISSEUR AB": "FR-MF-000000004",  # (and not the "-PR-" SRN,
                                          # which is a distributor)
    "FOURNISSEUR CD": "DE-MF-000000005",
    "FOURNISSEUR BF": "FR-MF-000000006",
    "FOURNISSEUR GE": "FR-MF-000000006",  # acquired by supplier BF
    "FOURNISSEUR E": "DE-MF-000000007",
}

CACHE = Path(__file__).resolve().parent / "wms_extracts" / ".cache_eudamed"


def _normaliser(texte) -> str:
    if texte is None:
        return ""
    texte = unicodedata.normalize("NFKD", str(texte))
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", texte.lower())


def ean_depuis_gtin(primary_di) -> str | None:
    """Converts a UDI-DI (GTIN-14) into an EAN-13.

    A GTIN-14 whose packaging indicator is 0 designates the sales unit:
    its EAN-13 is obtained by removing that zero. A non-zero indicator
    designates a pack, whose code must not be taken for the unit's.
    """
    if primary_di is None:
        return None
    chiffres = re.sub(r"\D", "", str(primary_di))
    if len(chiffres) == 14:
        return chiffres[1:] if chiffres[0] == "0" else None
    if len(chiffres) == 13:
        return chiffres
    return None


def _appeler(params: dict) -> tuple[int, list]:
    """One call to the API. Returns (total, items)."""
    reponse = requests.get(URL, params=params, headers=ENTETES, timeout=60)
    if "json" not in reponse.headers.get("content-type", ""):
        return 0, []
    donnees = reponse.json()
    return donnees.get("totalElements", 0), donnees.get("content") or []


def chercher(srn: str | None = None, trade_name: str | None = None,
             maximum: int = 1000) -> list[dict]:
    """Devices matching an SRN or a trade name.

    The total is checked: if it exceeds SEUIL_FILTRE_IGNORE, the API
    ignored the filter and is returning the whole database. We would
    rather return nothing than return anything at all.
    """
    if not srn and not trade_name:
        raise ValueError("An SRN or a trade name is required")

    base = {"languageIso2Code": "en", "size": TAILLE_PAGE,
            "pageSize": TAILLE_PAGE}
    if srn:
        base["srn"] = srn
    if trade_name:
        base["tradeName"] = trade_name

    total, premiers = _appeler({**base, "page": 0})
    if total > SEUIL_FILTRE_IGNORE:
        # Filter ignored: the answer means nothing
        return []

    resultats = list(premiers)
    pages = min((total + TAILLE_PAGE - 1) // TAILLE_PAGE,
                (maximum + TAILLE_PAGE - 1) // TAILLE_PAGE)
    for page in range(1, pages):
        time.sleep(PAUSE)
        _, suite = _appeler({**base, "page": page})
        if not suite:
            break
        resultats.extend(suite)

    return resultats[:maximum]


# Words too generic to identify a company on their own
MOTS_SOCIETE = {
    "GROUP", "GROUPE", "FRANCE", "SAS", "SARL", "SA", "GMBH", "AG", "BV",
    "NV", "LTD", "INC", "LP", "CO", "MEDICAL", "MEDIZIN", "SANTE", "HEALTH",
    "INDUSTRIES", "OPERATIONS", "SYSTEME", "SYSTEMES", "INTERNATIONAL",
}


def _mots_societe(nom) -> set:
    """Meaningful words of a company name."""
    if not nom:
        return set()
    texte = unicodedata.normalize("NFKD", str(nom).upper())
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    texte = re.sub(r"[^A-Z0-9 ]+", " ", texte)
    return {
        mot for mot in texte.split()
        if len(mot) >= 3 and mot not in MOTS_SOCIETE and not mot.isdigit()
    }


def _meme_societe(attendu, trouve) -> bool:
    """Is the manufacturer found really the one we were looking for?

    The comparison is made on WHOLE WORDS. A plain substring test would
    let "ALPHA" pass for "Alphamed Innovation", or "BETA" for
    "Betamedical": we would then inject another manufacturer's codes.
    """
    mots_attendus = _mots_societe(attendu)
    mots_trouves = _mots_societe(trouve)
    if not mots_attendus or not mots_trouves:
        return False
    return bool(mots_attendus & mots_trouves)


def trouver_srn(indices: list[str], fabricant: str | None = None
                ) -> tuple[str | None, str | None]:
    """Discovers a manufacturer's SRN from the names of its products.

    `indices` are known model names ("M200N", "Modele-D"). The first one
    that brings back a device from the right manufacturer yields its SRN.

    Returns (srn, manufacturer name as EUDAMED spells it).
    """

    for indice in indices:
        if not indice or len(indice) < 3:
            continue
        try:
            trouves = chercher(trade_name=indice, maximum=100)
        except requests.RequestException:
            continue
        time.sleep(PAUSE)

        for dispositif in trouves:
            nom = dispositif.get("manufacturerName") or ""
            srn = dispositif.get("manufacturerSrn")
            if not srn:
                continue
            if fabricant is None:
                return srn, nom
            # "Alpha Group" must answer to "ALPHA", but "Alphamed
            # Innovation" must not answer to "ALPHA".
            if _meme_societe(fabricant, nom):
                return srn, nom

    return None, None


def catalogue_fabricant(fournisseur: str, indices: list[str] | None = None,
                        utiliser_cache: bool = True) -> list[dict]:
    """Every device declared by one manufacturer.

    The SRN is looked up in SRN_CONNUS, then discovered from the product
    names supplied. The result is cached: the database moves little and
    the API is slow.
    """
    cle = _normaliser(fournisseur)
    CACHE.mkdir(parents=True, exist_ok=True)
    fichier = CACHE / f"{cle}.json"

    if utiliser_cache and fichier.exists():
        return json.loads(fichier.read_text(encoding="utf-8"))

    srn = SRN_CONNUS.get(fournisseur.upper())
    if srn is None and indices:
        srn, nom = trouver_srn(indices, fabricant=fournisseur)
        if srn:
            print(f"    SRN discovered for {fournisseur}: {srn} ({nom})")

    if not srn:
        # The failure is cached as well: without that, every run would
        # start a fresh series of searches for a supplier that has no
        # SRN, a distributor for instance.
        fichier.write_text("[]", encoding="utf-8")
        return []

    dispositifs = chercher(srn=srn, maximum=5000)
    fichier.write_text(
        json.dumps(dispositifs, ensure_ascii=False), encoding="utf-8"
    )
    return dispositifs


def indexer(dispositifs: list[dict]) -> dict:
    """Index of the devices by reference and by trade name.

    The EUDAMED reference sometimes carries a suffix ("REF12345 and
    suffix") that designates a family of variants: we index the stem as
    well.
    """
    par_reference: dict[str, dict] = {}
    par_nom: dict[str, list] = {}

    for dispositif in dispositifs:
        ean = ean_depuis_gtin(dispositif.get("primaryDi"))
        if not ean:
            continue
        enrichi = {**dispositif, "ean": ean}

        reference = dispositif.get("reference")
        if reference:
            for morceau in re.split(r"\s+and\s+|\s*[,;/]\s*", str(reference)):
                cle = _normaliser(morceau)
                if len(cle) >= 4 and cle not in par_reference:
                    par_reference[cle] = enrichi

        nom = _normaliser(dispositif.get("tradeName"))
        if nom:
            par_nom.setdefault(nom, []).append(enrichi)

    return {"par_reference": par_reference, "par_nom": par_nom}
