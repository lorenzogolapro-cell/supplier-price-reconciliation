# -*- coding: utf-8 -*-
"""
Fige le perimetre en LISTE, une fois pour toutes.

    python figer_perimetre.py            écrit la liste si elle n'existe pas
    python figer_perimetre.py --refaire  la réécrit (à n'utiliser que sur
                                         décision explicite)

Sortie : perimetre/perimetre_v1_1.csv
         perimetre/perimetre_v1_1.json   (manifeste)

Pourquoi une liste et pas un recalcul
--------------------------------------
Jusqu'ici chaque script recalculait le perimetre depuis les extraits.
Trois problemes, et le troisieme est le pire :

  1. c'est lent — 23 698 articles et 23 684 lignes de commande relues a
     chaque fois ;
  2. ca dépend de fichiers qui bougent — un nouvel export du WMS ou une
     ligne de commande de plus, et le perimetre change ;
  3. DEUX SESSIONS TRAVAILLENT EN PARALLELE sur ce perimetre. Si chacune
     le recalcule, rien ne garantit qu'elles parlent des memes articles.
     Un perimetre « figé » qui se recalcule n'est pas figé.

La liste est donc écrite une fois, avec l'empreinte des fichiers qui
l'ont produite. Toute session la LIT, aucune ne la recalcule.

Le garde-fou
------------
`VOLUME_ATTENDU` vaut 5 082 (v1.2 ; 4 731 etait v1.1, avant l'entree des
351 references du catalogue pro). Si le calcul rend autre chose, le script
s'arrete au lieu d'ecrire : c'est le signe qu'une source a bouge, et
cela demande une decision, pas un ecrasement silencieux.
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
    """Somme de controle d'une source, pour que la liste soit tracable."""
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
    """Le perimetre, calcule une derniere fois avant d'etre fige."""
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

    # Troisieme porte : le catalogue pro. On lit la liste figee du
    # catalogue, jamais le fichier source — les deux chantiers doivent
    # voir la meme population. Seul l'ENSEMBLE des references nous
    # interesse ici, pas sa colonne « au perimetre », qui elle depend
    # de la version du perimetre et serait circulaire.
    per["CATALOGUE_PRO"] = False
    if CATALOGUE_PRO_VAUT_ENTREE:
        if not CATALOGUE_PRO.exists():
            raise SystemExit(
                f"Liste du catalogue pro absente : {CATALOGUE_PRO}\n"
                f"La produire d'abord :  python figer_catalogue_pro.py")
        cat = pd.read_csv(CATALOGUE_PRO, dtype=str)
        refs = {normaliser_reference(v) for v in cat["Code article"]}
        per["CATALOGUE_PRO"] = per["_ka"].isin(refs - {None}).astype(bool)

    activite = per["VENTE_12M"] | per["ACHAT_12M"] | per["CATALOGUE_PRO"]
    # Les motifs d'ecartement passent TOUJOURS en premier : le catalogue
    # pro ouvre une porte, il ne leve pas un ecartement.
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
        print(f"\nLa liste existe déjà : {LISTE.name}")
        print(f"  {len(deja)} articles")
        print(f"  modifiée le "
              f"{datetime.fromtimestamp(LISTE.stat().st_mtime):%d/%m/%Y %H:%M}")
        print("\nElle n'est PAS recalculée : c'est tout l'intérêt.")
        print("Pour la refaire malgré tout : python figer_perimetre.py --refaire")
        return

    print(resume())
    print("\nCalcul du périmètre, une dernière fois avant de le figer...\n")
    liste, meta = calculer()

    print(f"  articles calculés .......... {len(liste)}")
    print(f"  volume attendu ............. {VOLUME_ATTENDU}")
    if len(liste) != VOLUME_ATTENDU:
        print(f"\n  ARRÊT — écart de {len(liste) - VOLUME_ATTENDU:+d} article(s).")
        print("  Une source a bougé. Rien n'est écrit : c'est une décision")
        print("  à prendre, pas un écrasement à faire en silence.")
        for s in meta["sources"]:
            print(f"      {s['fichier']:<40}{s['modifié']}  {s['sha256']}")
        raise SystemExit(1)
    print("  OK — le volume correspond.\n")

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
    print("\nÀ partir de maintenant : les sessions LISENT cette liste.")
    print("Voir perimetre_liste.py pour la charger en une ligne.")


if __name__ == "__main__":
    main()
