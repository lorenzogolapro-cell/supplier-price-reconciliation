# -*- coding: utf-8 -*-
"""
Purchase-price workstream: the single entry point.

    python run_pa.py                 the full run
    python run_pa.py --liste         prints the sequence, runs nothing
    python run_pa.py --sec           dry run: says what would be launched
    python run_pa.py --depuis pa_fusionner
    python run_pa.py --seulement pa_completer

WHY THIS FILE EXISTS
    The order of the scripts only ever lived in the documentation, so it
    kept getting lost. A skipped step does not show: the views happily
    regenerate from yesterday's pa_etat_par_article.xlsx, and the number
    that gets published becomes wrong without a single error being raised.

WHAT IT GUARANTEES
    1. The order. Each step reads what the previous one wrote.
    2. Stopping at the first failure. Carrying on after a failed step
       means propagating incomplete data all the way to the price
       owner's workbook.
    3. The start-up checks: the scope must load, and pa_completer.py
       must not land on a partial run.
    4. A summary: duration, return code and the files actually written.

WHAT IT DOES NOT DO
    It writes nothing itself and never touches any of the price owner's
    workbooks.
"""

from __future__ import annotations

import argparse
import subprocess
import sys
import time
from datetime import datetime
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

SORTIE_ACHATS = RACINE / "sortie" / "3-achats"


class Etape:
    """One step of the run.

    `bloquante` says whether a non-zero return code stops everything. The
    reconciliation is deliberately NON blocking here: it exits with 1 as
    soon as a reading divergence remains, which is exactly its job, but
    that finding must not prevent the read-only views from being
    regenerated. It prevents SHIPPING, and that is checked at shipping
    time, not here.
    """

    def __init__(self, cle, commande, role, bloquante=True):
        self.cle = cle
        self.commande = commande
        self.role = role
        self.bloquante = bloquante
        self.code = None
        self.duree = 0.0
        self.etat = "not run"


def enchainement() -> list[Etape]:
    """The documented running order, then the views derived from it."""
    etapes = [
        Etape("pa_completer", ["pa_completer.py", "--tous"],
              "sweep of the 154 suppliers, ~15 min"),
        Etape("pa_fusionner", ["pa_fusionner.py"],
              "pours the proposals into the price owner's workbook"),
        Etape("pa_etat_par_article", ["pa_etat_par_article.py"],
              "one row per item + reason, the base of every view"),
        # The check runs BEFORE the views: if the two pipelines are not
        # reading the same thing, the views are pretty and wrong.
        Etape("reconcilier", ["pa_reconcilier_responsable.py"],
              "our reading against the price owner's", bloquante=False),
    ]
    return etapes


# --------------------------------------------------------------------------
# Start-up checks


def perimetre_se_charge() -> bool:
    """The frozen scope must load, and the count must add up.

    We do not assume it: the module refuses to go on if the volume does
    not match the manifest, and that refusal is precisely what we want to
    trigger here rather than halfway through the sweep.
    """
    try:
        import perimetre_liste
        liste = perimetre_liste.charger()
    except Exception as err:
        print(f"  SCOPE FAILED: {type(err).__name__} - {err}")
        return False
    print(f"  scope: {len(liste)} items - {perimetre_liste.entete()}")
    return True


def passe_partielle() -> int:
    """How many pa_completer outputs are already present.

    `pa_completer.py --tous` SKIPS any supplier that already has an
    output file. Relaunching without archiving therefore replays nothing,
    and the run looks complete when in fact it recomputed nothing.
    """
    if not SORTIE_ACHATS.exists():
        return 0
    return len(list(SORTIE_ACHATS.glob("pa_completer_*.xlsx")))


def archiver_sorties() -> Path | None:
    """Moves the pa_completer_*.xlsx files into a dated subfolder.

    We archive, we do not delete: an output is the trace of a price list
    being read on a given date, and it is what lets us understand after
    the fact where a price came from.
    """
    fichiers = list(SORTIE_ACHATS.glob("pa_completer_*.xlsx"))
    if not fichiers:
        return None
    dest = SORTIE_ACHATS / f"_archive_{datetime.now():%Y-%m-%d}"
    dest.mkdir(parents=True, exist_ok=True)
    for f in fichiers:
        cible = dest / f.name
        if cible.exists():
            cible = dest / f"{f.stem} ({datetime.now():%H%M%S}){f.suffix}"
        f.rename(cible)
    print(f"  {len(fichiers)} outputs archived in {dest.name}/")
    return dest


# --------------------------------------------------------------------------


def lancer(etape: Etape, sec: bool) -> int:
    entete = f"  {etape.cle}  -  {etape.role}"
    print()
    print("-" * 74)
    print(entete)
    print(f"  $ python {' '.join(etape.commande)}")
    print("-" * 74)
    if sec:
        etape.etat = "simulated"
        return 0

    depart = time.time()
    # No capture: the sweep runs for a quarter of an hour and its progress
    # has to stay visible. We only collect the return code.
    resultat = subprocess.run([sys.executable] + etape.commande, cwd=str(RACINE))
    etape.duree = time.time() - depart
    etape.code = resultat.returncode
    etape.etat = "OK" if resultat.returncode == 0 else f"FAILED ({resultat.returncode})"
    return resultat.returncode


def recapituler(etapes: list[Etape], depart: float, arretee: Etape | None) -> None:
    print()
    print("=" * 74)
    print("SUMMARY")
    print("=" * 74)
    for e in etapes:
        duree = f"{e.duree:6.1f}s" if e.duree else "      -"
        print(f"  {e.etat:<14} {duree}  {e.cle}")
    print()
    print(f"  total duration: {time.time() - depart:.0f}s")

    if arretee is not None:
        print()
        print(f"  STOPPED at step '{arretee.cle}'. The following steps were")
        print("  not launched: the downstream outputs date from the previous")
        print("  run and must not be published.")
        return

    # What has just been written, so we know what to look at.
    recentes = [f for f in SORTIE_ACHATS.glob("*.xlsx")
                if f.stat().st_mtime > depart]
    if recentes:
        print()
        print("  written during this run:")
        for f in sorted(recentes, key=lambda p: p.stat().st_mtime):
            print(f"    {f.stat().st_size / 1024:>7.0f} KB  {f.name}")
    completer = len(list(SORTIE_ACHATS.glob("pa_completer_*.xlsx")))
    print()
    print(f"  {completer} suppliers swept by pa_completer")

    bloquant = [e for e in etapes if e.code not in (0, None)]
    if bloquant:
        print()
        print("  Steps that failed NON blocking, to look at before shipping:")
        for e in bloquant:
            print(f"    {e.cle} (code {e.code})")


def main() -> int:
    parseur = argparse.ArgumentParser(
        description="Runs the purchase-price workstream in order.")
    parseur.add_argument("--liste", action="store_true",
                         help="prints the sequence and exits")
    parseur.add_argument("--sec", action="store_true",
                         help="dry run, launches nothing")
    parseur.add_argument("--depuis", metavar="ETAPE",
                         help="resumes at this step")
    parseur.add_argument("--seulement", metavar="ETAPE",
                         help="runs this step only")
    parseur.add_argument("--archiver", action="store_true",
                         help="archive pa_completer_*.xlsx before the run")
    parseur.add_argument("--garder-sorties", action="store_true",
                         help="accepts a partial run without archiving")
    args = parseur.parse_args()

    etapes = enchainement()
    cles = [e.cle for e in etapes]

    if args.liste:
        print("Purchase-price workstream sequence:")
        for i, e in enumerate(etapes, 1):
            marque = " " if e.bloquante else "~"
            print(f" {marque}{i:>2}. {e.cle:<20} {e.role}")
        print()
        print(" ~ = a failure does not stop the run")
        return 0

    if args.seulement:
        if args.seulement not in cles:
            print(f"Unknown step: {args.seulement}. Known: {', '.join(cles)}")
            return 2
        etapes = [e for e in etapes if e.cle == args.seulement]
    elif args.depuis:
        if args.depuis not in cles:
            print(f"Unknown step: {args.depuis}. Known: {', '.join(cles)}")
            return 2
        etapes = etapes[cles.index(args.depuis):]

    print("=" * 74)
    print(f"PURCHASE-PRICE WORKSTREAM - run {datetime.now():%Y-%m-%d %H:%M}")
    print("=" * 74)
    print()
    print("Start-up checks")

    if not perimetre_se_charge():
        print()
        print("  The scope does not load: nothing can be launched.")
        return 1

    # A full run needs a clean slate, otherwise --tous replays nothing.
    rejoue_completer = any(e.cle == "pa_completer" for e in etapes)
    deja = passe_partielle()
    if rejoue_completer and deja:
        if args.archiver:
            archiver_sorties()
        elif args.garder_sorties:
            print(f"  {deja} pa_completer outputs kept: the sweep will skip "
                  f"those suppliers (resume, not a full run)")
        else:
            print(f"  {deja} pa_completer_*.xlsx files are already present.")
            print("  `--tous` skips suppliers already output, so the run")
            print("  would replay nothing. Relaunch with --archiver to move")
            print("  them aside, or --garder-sorties to resume.")
            return 1
    elif rejoue_completer:
        print("  no pending pa_completer output - full run")

    depart = time.time()
    arretee = None
    for etape in etapes:
        code = lancer(etape, args.sec)
        if code != 0 and etape.bloquante:
            arretee = etape
            break
        if code != 0:
            print(f"\n  ! {etape.cle} exits {code} - non blocking, continuing")

    recapituler(etapes, depart, arretee)
    return 1 if arretee else 0


if __name__ == "__main__":
    sys.exit(main())
