# -*- coding: utf-8 -*-
"""
Queries EUDAMED reference by reference, in the background.

    python eudamed_par_reference.py            resumes where it stopped
    python eudamed_par_reference.py --etat     where do we stand?
    python eudamed_par_reference.py --refaire  starts over

Output: wms_extracts/.cache_eudamed_ref/resultats.json
        sortie/2-chantier-ean/ean_eudamed_reference.xlsx

Why by reference
----------------
The whole previous collection went through the manufacturer's SRN, which
has to be known, and we only have eight of them. But the API also
accepts a `reference` filter, checked as an exact match: we can
therefore query any supplier at all, with no SRN, on the references we
are missing.

Throttling, and how we live with it
-----------------------------------
EUDAMED throttles silently: past roughly three calls a minute, it
answers HTTP 200 with zero results instead of a 429. A failure is
therefore indistinguishable from an absence, and that is what froze 248
empty caches we will never replay.

Two precautions:
  - a one minute pause between two calls, the rate at which no failure
    has been observed;
  - an empty result is NOT taken as final: the reference is retried up
    to REESSAIS times, at different moments. Only after that do we
    conclude it is absent.

Resuming after an interruption
------------------------------
The state is saved after every call. A power cut costs at most one
reference. The script can be relaunched as many times as needed: it
picks up where it left off.
"""

from __future__ import annotations

import json
import random
import re
import sys
import time
from datetime import datetime
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import pandas as pd
import requests

import eudamed_fiabilite  # noqa: E402
import mise_en_forme  # noqa: E402
from extracteurs.base import chemin_lisible, ean_est_valide  # noqa: E402
from main import SORTIE_CHANTIER, WMS_EXTRACTS  # noqa: E402

URL = "https://ec.europa.eu/tools/eudamed/api/devices/udiDiData"
ENTETES = {
    "Accept": "application/json",
    "Accept-Language": "en",
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/120.0 Safari/537.36"),
}

CACHE = WMS_EXTRACTS / ".cache_eudamed_ref"
ETAT = CACHE / "resultats.json"

# One minute between two calls: the rate at which no silent failure has
# been observed. Below that, the API answers 200 with zero results.
PAUSE = 60
# We vary it slightly so as not to hit on the exact second
JITTER = 12

# How many times to insist on a reference that came back empty.
#
# The setting was 3, out of fear of silent throttling. Measured on 04/09
# over 1 195 references, the hypothesis does not hold:
#
#     1st attempt  494 references   318 found   64 %
#     2nd attempt  620 references    15 found   2.4 %
#     3rd attempt   45 references     2 found   4.4 %
#
# An empty first call is therefore a REAL absence, not a disguised
# refusal: at one call a minute, EUDAMED does not slow us down.
# Insisting cost thirteen hours for about twenty codes, while thousands
# of references had never been queried even once and had, for their
# part, two chances out of three of succeeding.
#
# So we exhaust the first attempts first. A second pass will be decided
# afterwards, on the confirmed absences, once EUDAMED has filled up: the
# regulation applies in stages.
REESSAIS = 1

TIMEOUT = 90


def _empecher_la_veille() -> bool:
    """Asks Windows not to fall asleep while we are working.

    The collection runs for days. The machine, however, goes to sleep
    after a few dozen minutes of inactivity, and a collection that uses
    neither keyboard nor mouse counts as inactivity. We then lose the
    whole night, and we have to wait for the next day's logon before the
    launcher resumes.

    `SetThreadExecutionState` settles this without any special rights:
    the process declares that the system must stay awake. We do NOT ask
    for ES_DISPLAY_REQUIRED: the screen can go off, it is of no use here,
    and leaving it on would wear the panel for nothing.

    The effect ends by itself when the process terminates: nothing to
    undo, nothing left hanging if the collection is killed.
    """
    if sys.platform != "win32":
        return False
    try:
        import ctypes
        ES_CONTINUOUS = 0x80000000
        ES_SYSTEM_REQUIRED = 0x00000001
        resultat = ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
        return bool(resultat)
    except Exception:
        # A domain policy may refuse it: that is no reason to stop the
        # collection, only to say so.
        return False


def _charger() -> dict:
    if ETAT.exists():
        try:
            return json.loads(ETAT.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            pass
    return {}


def _enregistrer(etat: dict) -> None:
    CACHE.mkdir(parents=True, exist_ok=True)
    provisoire = ETAT.with_suffix(".tmp")
    provisoire.write_text(json.dumps(etat, ensure_ascii=False),
                          encoding="utf-8")
    provisoire.replace(ETAT)


COLONNES_ARTICLE = ["Code article", "Réf. fournisseur", "Désignation",
                    "Fournisseur", "Actif"]


def _tous_les_articles() -> pd.DataFrame:
    """Every article carrying a supplier reference.

    Two uses, and they must not be confused: this table is here to NAME
    the articles when the workbook of findings is written. `_a_chercher`
    then derives from it what is left to query.
    """
    fichier = mise_en_forme.derniere_version(SORTIE_CHANTIER,
                                             "EAN distributeur.xlsx")
    if fichier is None:
        raise SystemExit("Run first: python ean_global.py")
    df = pd.read_excel(chemin_lisible(fichier), sheet_name="Codes EAN",
                       skiprows=3, dtype=str)
    return df[df["Code article"].notna() & df["Réf. fournisseur"].notna()]


# A reference EUDAMED can work with: digits, spaces and separating dots.
# "25 105 00" and "72.700" pass, "ALPHA2_MO_T15" does not.
REFERENCE_NUMERIQUE = re.compile(r"[\d\s.]{4,}")


def _interrogeable(article: dict) -> bool:
    """Does this reference stand a chance of being known to EUDAMED?

    The question only arises for references taken from the WMS
    "Référence fabricant" column, opened as a fallback when the
    supplier's one is missing. That column mixes genuine catalogue
    references with in-house configuration codes.

    Two pilots of 25 real calls measured it:

        numeric references         12 hits out of 38   32 %
        alphanumeric references     0 hits out of 12    0 %

    "KIT_ASM_RLV01", "GAMME MS-BASE", "ALPHA2_MO_T15": no manufacturer
    declares a device under that kind of string. Querying them would cost
    27 hours for nothing.

    The SUPPLIER reference, on the other hand, passes unconditionally: it
    returns something 27.6 % of the time whatever its shape, measured
    over 3 658 calls.

    One more condition applies since 11/09/2026, and it holds for BOTH
    columns: the reference must be at least
    `eudamed_fiabilite.LONGUEUR_MINIMALE` characters long. "Returning
    27.6 %" did not mean "returning the right thing": the measurement
    counted the calls that brought back SOMETHING, not those that brought
    back the right device. Confronted with the supplier's GS1 prefix,
    codes drawn from a short reference fall to 15 % agreement, against
    96 % beyond nine characters. Querying them does not only cost time:
    it produces falsehoods that then have to be undone.
    """
    reference = str(article.get("Réf. fournisseur") or "").strip()
    if not eudamed_fiabilite.reference_fiable(reference):
        return False
    if article.get("Origine référence") != "fabricant":
        return True
    return bool(REFERENCE_NUMERIQUE.fullmatch(reference))


def _a_chercher(tous: pd.DataFrame | None = None) -> list[dict]:
    """Our gaps that have a usable reference, active articles first."""
    df = _tous_les_articles() if tous is None else tous
    df = df[df["Code EAN"].isna()]
    # The articles that actually move go first: if the collection stops
    # part way, they are the ones we will have documented.
    df = df.sort_values("Actif", ascending=False)

    colonnes = COLONNES_ARTICLE + (
        ["Origine référence"] if "Origine référence" in df.columns else [])
    articles = df[colonnes].to_dict("records")
    retenus = [a for a in articles if _interrogeable(a)]
    ecartes = len(articles) - len(retenus)
    if ecartes:
        print(f"{ecartes} alphanumeric manufacturer references discarded, "
              f"measured at a 0 % hit rate")
    return retenus


def _interroger(reference: str) -> tuple[str | None, str]:
    """(EAN, status) for one reference. status: trouve / absent / echec."""
    try:
        reponse = requests.get(
            URL, headers=ENTETES, timeout=TIMEOUT,
            params={"languageIso2Code": "en", "size": 10, "pageSize": 10,
                    "page": 0, "reference": reference})
    except requests.RequestException:
        return None, "echec"

    if reponse.status_code != 200:
        return None, "echec"
    if "json" not in reponse.headers.get("content-type", ""):
        return None, "echec"

    donnees = reponse.json()
    total = donnees.get("totalElements", 0)
    # An outsized total signals an ignored filter, not a result
    if total > 100_000:
        return None, "echec"
    if not total:
        return None, "absent"

    for element in donnees.get("content") or []:
        brut = str(element.get("primaryDi") or "")
        if len(brut) == 14 and brut[0] == "0":
            candidat = brut[1:]
        elif len(brut) == 13:
            candidat = brut
        else:
            continue
        if ean_est_valide(candidat):
            return candidat, "trouve"
    # Devices do exist, but none of them has a sales unit code
    return None, "absent"


def _ecrire_resultats(etat: dict, tous: pd.DataFrame) -> None:
    """Workbook of the codes found, ready to be picked up.

    Written from ALL the articles, never from the list of what is left to
    look for. The distinction cost 221 codes: that list only holds the
    articles WITHOUT a code, and an article EUDAMED has just documented
    is no longer part of it on the next pass. Its code therefore
    disappeared from the workbook, `ean_global` no longer saw it, and the
    article found itself without a code again, then back in the queue.
    The collection was erasing its own findings.

    One reference can name several of our articles: the code then holds
    for each of them, and each one gets its row.
    """
    par_reference: dict[str, list[dict]] = {}
    for article in tous[COLONNES_ARTICLE].to_dict("records"):
        par_reference.setdefault(article["Réf. fournisseur"], []).append(
            article)

    lignes = []
    for reference, valeur in etat.items():
        if not valeur.get("ean"):
            continue
        for article in par_reference.get(reference, []):
            lignes.append({
                "Code article": article["Code article"],
                "Référence fournisseur": reference,
                "Désignation": article["Désignation"],
                "Fournisseur": article["Fournisseur"],
                "Actif": article["Actif"],
                "Code EAN": valeur["ean"],
                "Source": "eudamed-référence",
                "Relevé le": valeur.get("date", ""),
            })
    if not lignes:
        return

    df = pd.DataFrame(lignes)
    SORTIE_CHANTIER.mkdir(parents=True, exist_ok=True)
    chemin = mise_en_forme.chemin_ecriture(
        SORTIE_CHANTIER / "ean_eudamed_reference.xlsx")
    with pd.ExcelWriter(chemin, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Codes EAN", index=False, startrow=3)
    mise_en_forme.formater(
        chemin,
        options={"Codes EAN": {
            "ligne_entete": 4,
            "titre": "Codes EAN relevés sur EUDAMED, par référence",
            "sous_titre": (
                f"{len(df)} codes | correspondance exacte sur la référence "
                f"fabricant, donc sûre"
            ),
        }},
    )


def etat_courant() -> None:
    etat = _charger()
    articles = _a_chercher(_tous_les_articles())
    trouves = sum(1 for v in etat.values() if v.get("ean"))
    absents = sum(1 for v in etat.values()
                  if v.get("statut") == "absent"
                  and v.get("essais", 0) >= REESSAIS)
    a_reessayer = sum(1 for v in etat.values()
                      if not v.get("ean") and v.get("essais", 0) < REESSAIS)
    restants = [a for a in articles
                if etat.get(a["Réf. fournisseur"], {}).get("essais", 0)
                < REESSAIS and not etat.get(a["Réf. fournisseur"],
                                            {}).get("ean")]
    print(f"  {len(articles)} references to document")
    print(f"  {trouves} found | {absents} absent (confirmed) | "
          f"{a_reessayer} to retry")
    print(f"  {len(restants)} calls left, i.e. roughly "
          f"{len(restants) * PAUSE / 3600:.0f} h")


def main() -> None:
    if "--refaire" in sys.argv and ETAT.exists():
        ETAT.unlink()
    if "--etat" in sys.argv:
        etat_courant()
        return

    if _empecher_la_veille():
        print("system sleep disabled for the duration of the collection "
              "(the screen may go off)")
    else:
        print("! sleep could not be prevented: a night of collection may "
              "be lost if the machine falls asleep")

    etat = _charger()
    tous = _tous_les_articles()
    articles = _a_chercher(tous)

    # What is left: never tried, or tried without success fewer than
    # REESSAIS times.
    #
    # ONCE PER REFERENCE ONLY. Several of our articles often share the
    # same supplier reference: variants created separately, or duplicates
    # in the database. Without this deduplication, each of them triggered
    # its own call for the same question: one reference was queried 29
    # times, another 10, and the attempt counter kept climbing without
    # anything new being asked. That is wasted call time, and it is also
    # what made it look as though we were hammering exhausted references.
    restants, vues = [], set()
    for article in articles:
        reference = article["Réf. fournisseur"]
        if reference in vues:
            continue
        vues.add(reference)
        connu = etat.get(reference, {})
        if connu.get("ean") or connu.get("essais", 0) >= REESSAIS:
            continue
        restants.append(article)

    trouves = sum(1 for v in etat.values() if v.get("ean"))
    print(f"{len(articles)} articles | {len(vues)} distinct references | "
          f"{trouves} already found")
    print(f"{len(restants)} to query, ~{len(restants) * PAUSE / 3600:.0f} h "
          f"at one call a minute\n")

    depuis_ecriture = 0
    for rang, article in enumerate(restants, start=1):
        reference = article["Réf. fournisseur"]
        ean, statut = _interroger(reference)

        entree = etat.setdefault(reference, {"essais": 0})
        entree["essais"] = entree.get("essais", 0) + 1
        entree["statut"] = statut
        entree["date"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        if ean:
            entree["ean"] = ean
            trouves += 1

        marque = {"trouve": "OK  ", "absent": "  --", "echec": " !! "}[statut]
        print(f"  [{rang:>5}/{len(restants)}] {marque} "
              f"{str(reference)[:20]:<20} {ean or ''}  "
              f"({trouves} found)", flush=True)

        _enregistrer(etat)
        depuis_ecriture += 1
        if depuis_ecriture >= 25:
            _ecrire_resultats(etat, tous)
            depuis_ecriture = 0

        if rang < len(restants):
            time.sleep(PAUSE + random.uniform(0, JITTER))

    _ecrire_resultats(etat, tous)
    print(f"\n{trouves} EAN codes found on EUDAMED")


if __name__ == "__main__":
    main()
