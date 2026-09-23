# -*- coding: utf-8 -*-
"""
Chantier PA — le point d'entree unique.

    python run_pa.py                 la passe complete
    python run_pa.py --liste         affiche l'enchainement, ne lance rien
    python run_pa.py --sec           simule : dit ce qui serait lance
    python run_pa.py --depuis pa_fusionner
    python run_pa.py --seulement pa_completer

POURQUOI CE FICHIER
    L'ordre des scripts n'existait que dans la documentation, donc il se
    perdait. Une etape sautee ne se voit pas : les vues se regenerent sans
    broncher a partir d'un pa_etat_par_article.xlsx de la veille, et le
    chiffre publie devient faux sans qu'aucune erreur ne soit levee.

CE QU'IL GARANTIT
    1. L'ordre. Chaque etape lit ce que la precedente a ecrit.
    2. L'arret au premier echec. Continuer apres une etape ratee, c'est
       propager une donnee incomplete jusqu'au classeur du responsable
       des prix.
    3. Le controle de depart : le perimetre doit se charger, et
       pa_completer.py ne doit pas retomber sur une passe partielle.
    4. Un recapitulatif : duree, code retour et fichiers reellement ecrits.

CE QU'IL NE FAIT PAS
    Il n'ecrit rien lui-meme et ne touche a aucun classeur du responsable
    des prix.
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
    """Une etape de la passe.

    `bloquante` dit si un code retour non nul arrete tout. La reconciliation
    est volontairement NON bloquante ici : elle sort en 1 des qu'il reste une
    divergence de lecture, ce qui est son role, mais ce constat ne doit pas
    empecher de regenerer les vues de lecture. Il empeche d'EXPEDIER, ce
    qui se verifie au moment d'expedier, pas ici.
    """

    def __init__(self, cle, commande, role, bloquante=True):
        self.cle = cle
        self.commande = commande
        self.role = role
        self.bloquante = bloquante
        self.code = None
        self.duree = 0.0
        self.etat = "non lancée"


def enchainement() -> list[Etape]:
    """L'ordre de CLAUDE.md §3, puis les vues qui en derivent."""
    etapes = [
        Etape("pa_completer", ["pa_completer.py", "--tous"],
              "balayage des 154 fournisseurs, ~15 min"),
        Etape("pa_fusionner", ["pa_fusionner.py"],
              "verse les propositions dans le classeur du responsable"),
        Etape("pa_etat_par_article", ["pa_etat_par_article.py"],
              "une ligne par article + motif — socle de toutes les vues"),
        # Le controle passe AVANT les vues : si les deux pipelines ne lisent
        # pas la meme chose, les vues sont belles et fausses.
        Etape("reconcilier", ["pa_reconcilier_responsable.py"],
              "notre lecture contre celle du responsable", bloquante=False),
    ]
    return etapes


# --------------------------------------------------------------------------
# Controles de depart


def perimetre_se_charge() -> bool:
    """Le perimetre figé doit se lire, et le compte doit tomber.

    On ne le suppose pas : le module refuse de continuer si le volume ne
    correspond pas au manifeste, et c'est precisement ce refus qu'on veut
    declencher ici plutot qu'au milieu du balayage.
    """
    try:
        import perimetre_liste
        liste = perimetre_liste.charger()
    except Exception as err:
        print(f"  ECHEC perimetre : {type(err).__name__} — {err}")
        return False
    print(f"  perimetre : {len(liste)} articles — {perimetre_liste.entete()}")
    return True


def passe_partielle() -> int:
    """Nombre de sorties pa_completer deja presentes.

    `pa_completer.py --tous` SAUTE tout fournisseur qui a deja un fichier de
    sortie. Relancer sans avoir archive ne rejoue donc rien, et la passe
    passe pour complete alors qu'elle n'a rien recalcule.
    """
    if not SORTIE_ACHATS.exists():
        return 0
    return len(list(SORTIE_ACHATS.glob("pa_completer_*.xlsx")))


def archiver_sorties() -> Path | None:
    """Deplace les pa_completer_*.xlsx dans un sous-dossier date.

    On archive, on ne supprime pas : une sortie est la trace d'une lecture
    de tarif a une date donnee, et c'est elle qui permet de comprendre
    apres coup d'ou venait un prix.
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
    print(f"  {len(fichiers)} sorties archivees dans {dest.name}/")
    return dest


# --------------------------------------------------------------------------


def lancer(etape: Etape, sec: bool) -> int:
    entete = f"  {etape.cle}  —  {etape.role}"
    print()
    print("-" * 74)
    print(entete)
    print(f"  $ python {' '.join(etape.commande)}")
    print("-" * 74)
    if sec:
        etape.etat = "simulée"
        return 0

    depart = time.time()
    # Pas de capture : le balayage dure un quart d'heure et son avancement
    # doit rester visible. On ne recupere que le code retour.
    resultat = subprocess.run([sys.executable] + etape.commande, cwd=str(RACINE))
    etape.duree = time.time() - depart
    etape.code = resultat.returncode
    etape.etat = "OK" if resultat.returncode == 0 else f"ECHEC ({resultat.returncode})"
    return resultat.returncode


def recapituler(etapes: list[Etape], depart: float, arretee: Etape | None) -> None:
    print()
    print("=" * 74)
    print("RECAPITULATIF")
    print("=" * 74)
    for e in etapes:
        duree = f"{e.duree:6.1f}s" if e.duree else "      -"
        print(f"  {e.etat:<14} {duree}  {e.cle}")
    print()
    print(f"  duree totale : {time.time() - depart:.0f}s")

    if arretee is not None:
        print()
        print(f"  ARRET a l'etape « {arretee.cle} ». Les etapes suivantes n'ont")
        print("  pas ete lancees : les sorties en aval datent de la passe")
        print("  precedente et ne doivent pas etre publiees.")
        return

    # Ce qui vient d'etre ecrit, pour qu'on sache quoi regarder.
    recentes = [f for f in SORTIE_ACHATS.glob("*.xlsx")
                if f.stat().st_mtime > depart]
    if recentes:
        print()
        print("  ecrits pendant cette passe :")
        for f in sorted(recentes, key=lambda p: p.stat().st_mtime):
            print(f"    {f.stat().st_size / 1024:>7.0f} Ko  {f.name}")
    completer = len(list(SORTIE_ACHATS.glob("pa_completer_*.xlsx")))
    print()
    print(f"  {completer} fournisseurs balayes par pa_completer")

    bloquant = [e for e in etapes if e.code not in (0, None)]
    if bloquant:
        print()
        print("  Etapes en echec NON bloquant, a regarder avant d'expedier :")
        for e in bloquant:
            print(f"    {e.cle} (code {e.code})")


def main() -> int:
    parseur = argparse.ArgumentParser(
        description="Enchaine le chantier PA dans l'ordre.")
    parseur.add_argument("--liste", action="store_true",
                         help="affiche l'enchainement et sort")
    parseur.add_argument("--sec", action="store_true",
                         help="simule sans rien lancer")
    parseur.add_argument("--depuis", metavar="ETAPE",
                         help="reprend a cette etape")
    parseur.add_argument("--seulement", metavar="ETAPE",
                         help="ne lance que cette etape")
    parseur.add_argument("--archiver", action="store_true",
                         help="archive les pa_completer_*.xlsx avant la passe")
    parseur.add_argument("--garder-sorties", action="store_true",
                         help="accepte une passe partielle sans archiver")
    args = parseur.parse_args()

    etapes = enchainement()
    cles = [e.cle for e in etapes]

    if args.liste:
        print("Enchainement du chantier PA :")
        for i, e in enumerate(etapes, 1):
            marque = " " if e.bloquante else "~"
            print(f" {marque}{i:>2}. {e.cle:<20} {e.role}")
        print()
        print(" ~ = un echec n'arrete pas la passe")
        return 0

    if args.seulement:
        if args.seulement not in cles:
            print(f"Etape inconnue : {args.seulement}. Connues : {', '.join(cles)}")
            return 2
        etapes = [e for e in etapes if e.cle == args.seulement]
    elif args.depuis:
        if args.depuis not in cles:
            print(f"Etape inconnue : {args.depuis}. Connues : {', '.join(cles)}")
            return 2
        etapes = etapes[cles.index(args.depuis):]

    print("=" * 74)
    print(f"CHANTIER PA — passe du {datetime.now():%Y-%m-%d %H:%M}")
    print("=" * 74)
    print()
    print("Controles de depart")

    if not perimetre_se_charge():
        print()
        print("  Le perimetre ne se charge pas : rien ne peut etre lance.")
        return 1

    # Une passe complete exige de repartir a zero, sinon --tous ne rejoue rien.
    rejoue_completer = any(e.cle == "pa_completer" for e in etapes)
    deja = passe_partielle()
    if rejoue_completer and deja:
        if args.archiver:
            archiver_sorties()
        elif args.garder_sorties:
            print(f"  {deja} sorties pa_completer conservees : le balayage "
                  f"sautera ces fournisseurs (reprise, pas passe complete)")
        else:
            print(f"  {deja} fichiers pa_completer_*.xlsx sont deja presents.")
            print("  `--tous` saute les fournisseurs deja sortis : la passe ne")
            print("  rejouerait donc rien. Relancer avec --archiver pour les")
            print("  mettre de cote, ou --garder-sorties pour une reprise.")
            return 1
    elif rejoue_completer:
        print("  aucune sortie pa_completer en attente — passe complete")

    depart = time.time()
    arretee = None
    for etape in etapes:
        code = lancer(etape, args.sec)
        if code != 0 and etape.bloquante:
            arretee = etape
            break
        if code != 0:
            print(f"\n  ! {etape.cle} sort en {code} — non bloquant, on continue")

    recapituler(etapes, depart, arretee)
    return 1 if arretee else 0


if __name__ == "__main__":
    sys.exit(main())
