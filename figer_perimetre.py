# -*- coding: utf-8 -*-
"""
Freezes the scope into a LIST, once and for all.

    python figer_perimetre.py            writes the list if it does not
                                         exist yet
    python figer_perimetre.py --refaire  rewrites it (to be used only on
                                         an explicit decision)

Output: perimetre/perimetre_v1_1.csv
        perimetre/perimetre_v1_1.json   (manifest)

Why a list and not a recompute
------------------------------
Until now every script recomputed the scope from the extracts. Three
problems, and the third one is the worst:

  1. it is slow: 23,698 items and 23,684 order lines re-read every
     single time;
  2. it depends on files that move: one new WMS export or one extra
     order line, and the scope changes;
  3. TWO SESSIONS WORK IN PARALLEL on this scope. If each one
     recomputes it, nothing guarantees they are talking about the same
     items. A "frozen" scope that gets recomputed is not frozen.

So the list is written once, with the fingerprint of the files that
produced it. Every session READS it, none recomputes it.

The guardrail
-------------
`VOLUME_ATTENDU` is 5,082 (v1.2; 4,731 was v1.1, before the 351 pro
catalogue references came in). If the computation returns anything
else, the script stops instead of writing: it is the sign that a source
has moved, and that calls for a decision, not a silent overwrite.
"""

from __future__ import annotations

import hashlib
import json
import sys
import warnings
from datetime import datetime
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

warnings.filterwarnings("ignore")

import pandas as pd  # noqa: E402

from perimetre_comptage import ARRETE, ARTICLES  # noqa: E402
from perimetre_config import (  # noqa: E402
    ACHAT_VAUT_ENTREE,
    CATALOGUE_PRO_VAUT_ENTREE,
    FENETRE_VENTE_MOIS,
    PERIMETRE_VERSION,
    VOLUME_ATTENDU,
    resume,
)
from normalisation import normaliser_reference  # noqa: E402

DOSSIER = RACINE / "perimetre"
LISTE = DOSSIER / f"perimetre_v{PERIMETRE_VERSION.replace('.', '_')}.csv"
MANIFESTE = DOSSIER / f"perimetre_v{PERIMETRE_VERSION.replace('.', '_')}.json"
CATALOGUE_PRO = DOSSIER / "catalogue_pro.csv"

COLONNES = [
    "Réf. interne", "Code article", "Code déclinaison", "Désignation",
    "Fournisseur", "Type", "Famille", "VENTE_12M", "ACHAT_12M",
    "CATALOGUE_PRO", "PORTE",
]


def empreinte(chemin: Path) -> dict:
    """Checksum of a source, so the list stays traceable."""
    condense = hashlib.sha256()
    with chemin.open("rb") as flux:
        for bloc in iter(lambda: flux.read(1 << 20), b""):
            condense.update(bloc)
    horodatage = datetime.fromtimestamp(chemin.stat().st_mtime)
    return {
        "fichier": chemin.name,
        "sha256": condense.hexdigest()[:16],
        "modifié": horodatage.strftime("%Y-%m-%d %H:%M"),
        "octets": chemin.stat().st_size,
    }


def calculer() -> tuple[pd.DataFrame, dict]:
    """The scope, computed one last time before being frozen."""
    from achats_par_article import (COMMANDES, chemin_lisible,
                                    construire_table)
    from perimetre_completude import charger_sortie

    per = charger_sortie()
    cmd = pd.read_excel(chemin_lisible(COMMANDES), dtype=object)
    table, _ = construire_table(cmd)
    idx = table.set_index("Réf. Art.")

    per["ACHAT_12M"] = (per["_cle"].map(idx["STATUT_ACHAT"])
                        == "ACHAT_CONFIRME").astype(bool)
    per["VENTE_12M"] = per["TOURNE"].astype(bool)
    per["_ka"] = per["Code article"].map(normaliser_reference)

    # Third door: the pro catalogue. We read the frozen catalogue list,
    # never the source file: both workstreams have to see the same
    # population. Only the SET of references matters here, not its
    # "inside the scope" column, which depends on the scope version and
    # would be circular.
    per["CATALOGUE_PRO"] = False
    if CATALOGUE_PRO_VAUT_ENTREE:
        if not CATALOGUE_PRO.exists():
            raise SystemExit(
                f"Pro catalogue list missing: {CATALOGUE_PRO}\n"
                f"Produce it first:  python figer_catalogue_pro.py")
        cat = pd.read_csv(CATALOGUE_PRO, dtype=str)
        refs = {normaliser_reference(v) for v in cat["Code article"]}
        per["CATALOGUE_PRO"] = per["_ka"].isin(refs - {None}).astype(bool)

    activite = per["VENTE_12M"] | per["ACHAT_12M"] | per["CATALOGUE_PRO"]
    # The exclusion reasons ALWAYS come first: the pro catalogue opens a
    # door, it does not lift an exclusion.
    dedans = per[per["au_perimetre"] & activite].copy()

    def porte(v: bool, a: bool, c: bool) -> str:
        if v and a:
            return "vente + achat"
        if v:
            return "vente seule"
        if a:
            return "achat seul"
        return "catalogue pro"

    dedans["PORTE"] = [porte(v, a, c) for v, a, c in
                       zip(dedans["VENTE_12M"], dedans["ACHAT_12M"],
                           dedans["CATALOGUE_PRO"])]
    liste = pd.DataFrame({
        "Réf. interne": dedans["_cle"],
        "Code article": dedans["_ka"],
        "Code déclinaison": dedans["Code déclinaison"].map(
            normaliser_reference),
        "Désignation": dedans["Libellé déclinaison ^(1)"],
        "Fournisseur": dedans["Nom fournisseur"],
        "Type": dedans["Type"],
        "Famille": dedans["Libellé famille"],
        "VENTE_12M": dedans["VENTE_12M"],
        "ACHAT_12M": dedans["ACHAT_12M"],
        "CATALOGUE_PRO": dedans["CATALOGUE_PRO"],
        "PORTE": dedans["PORTE"],
    })[COLONNES].sort_values("Code article").reset_index(drop=True)

    sources = [empreinte(ARTICLES), empreinte(chemin_lisible(COMMANDES))]
    if CATALOGUE_PRO_VAUT_ENTREE:
        sources.append(empreinte(CATALOGUE_PRO))
    return liste, {"sources": sources}


def main() -> None:
    refaire = "--refaire" in sys.argv
    DOSSIER.mkdir(parents=True, exist_ok=True)

    if LISTE.exists() and not refaire:
        deja = pd.read_csv(LISTE, dtype=str)
        print(resume())
        print(f"\nThe list already exists: {LISTE.name}")
        print(f"  {len(deja)} items")
        print(f"  last modified on "
              f"{datetime.fromtimestamp(LISTE.stat().st_mtime):%d/%m/%Y %H:%M}")
        print("\nIt is NOT recomputed: that is the whole point.")
        print("To rebuild it anyway: python figer_perimetre.py --refaire")
        return

    print(resume())
    print("\nComputing the scope, one last time before freezing it...\n")
    liste, meta = calculer()

    print(f"  items computed ............. {len(liste)}")
    print(f"  expected volume ............ {VOLUME_ATTENDU}")
    if len(liste) != VOLUME_ATTENDU:
        print(f"\n  STOP - gap of {len(liste) - VOLUME_ATTENDU:+d} item(s).")
        print("  A source has moved. Nothing is written: this is a decision")
        print("  to take, not an overwrite to do quietly.")
        for s in meta["sources"]:
            print(f"      {s['fichier']:<40}{s['modifié']}  {s['sha256']}")
        raise SystemExit(1)
    print("  OK - the volume matches.\n")

    for v, n in liste["PORTE"].value_counts().items():
        print(f"      {n:>6}  {v}")

    liste.to_csv(LISTE, index=False, encoding="utf-8-sig")

    manifeste = {
        "version": PERIMETRE_VERSION,
        "figé_le": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "articles": len(liste),
        "règle": ("aucun motif écarté ET (vente 12 mois OU achat reçu "
                  "12 mois)"),
        "fenêtre_mois": FENETRE_VENTE_MOIS,
        "arrêté": ARRETE.strftime("%Y-%m-%d"),
        "sources": meta["sources"],
        "portes": {str(k): int(v)
                   for k, v in liste["PORTE"].value_counts().items()},
    }
    MANIFESTE.write_text(
        json.dumps(manifeste, ensure_ascii=False, indent=2), encoding="utf-8")

    print(f"\n  -> {LISTE}")
    print(f"  -> {MANIFESTE}")
    print("\nFrom now on: sessions READ this list.")
    print("See perimetre_liste.py to load it in one line.")


if __name__ == "__main__":
    main()
