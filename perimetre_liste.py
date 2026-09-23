# -*- coding: utf-8 -*-
"""
Le perimetre fige, en lecture seule.

C'est le point d'entree pour TOUTE session qui travaille sur le
perimetre — chantier PA, chantier EAN, ou autre.

    from perimetre_liste import charger, references

    liste = charger()            # DataFrame, 4 731 lignes
    refs = references()          # set de Références internes
    if ref in refs: ...

Pourquoi ne PAS recalculer
---------------------------
Le perimetre est fige. Le recalculer a chaque session, c'est risquer que
deux sessions travaillent sur deux populations differentes sans que
personne ne s'en apercoive — il suffit qu'un extrait ait ete regenere
entre les deux.

Ce module lit la liste ecrite une fois par `figer_perimetre.py`. Il ne
recalcule rien, jamais. Si la liste manque, il le dit et s'arrete plutot
que de la reconstituer a la volee.

Si le perimetre doit changer, c'est une decision : on incremente
PERIMETRE_VERSION et on relance `figer_perimetre.py --refaire`.
"""

from __future__ import annotations

import json
import sys
from functools import lru_cache
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import pandas as pd  # noqa: E402

from perimetre_config import PERIMETRE_VERSION, VOLUME_ATTENDU  # noqa: E402

DOSSIER = RACINE / "perimetre"
_SUFFIXE = PERIMETRE_VERSION.replace(".", "_")
LISTE = DOSSIER / f"perimetre_v{_SUFFIXE}.csv"
MANIFESTE = DOSSIER / f"perimetre_v{_SUFFIXE}.json"

_ABSENTE = (
    f"Le périmètre figé est introuvable : {LISTE}\n"
    f"Le produire avec :  python figer_perimetre.py\n"
    f"Ne PAS le recalculer dans votre script : deux sessions qui "
    f"recalculent peuvent diverger."
)


@lru_cache(maxsize=1)
def charger() -> pd.DataFrame:
    """Le périmètre figé. Lu une fois, mis en cache pour la session."""
    if not LISTE.exists():
        raise SystemExit(_ABSENTE)
    df = pd.read_csv(LISTE, dtype=str)
    for colonne in ("VENTE_12M", "ACHAT_12M"):
        if colonne in df.columns:
            df[colonne] = df[colonne].astype(str).str.lower() == "true"
    if len(df) != VOLUME_ATTENDU:
        raise SystemExit(
            f"Le périmètre figé compte {len(df)} articles, "
            f"{VOLUME_ATTENDU} attendus. La liste a été modifiée hors "
            f"figer_perimetre.py — vérifier avant d'aller plus loin.")
    return df


@lru_cache(maxsize=1)
def references() -> frozenset:
    """Les Références internes du périmètre, pour un test d'appartenance."""
    return frozenset(charger()["Réf. interne"].dropna())


@lru_cache(maxsize=1)
def codes_article() -> frozenset:
    """Les codes article du périmètre.

    À n'utiliser que pour un rapprochement INTERNE d'un export WMS à un
    autre. Vers toute source extérieure, la clé reste la Référence.
    """
    return frozenset(charger()["Code article"].dropna())


def manifeste() -> dict:
    """Version, date de figeage et empreinte des sources."""
    if not MANIFESTE.exists():
        return {}
    return json.loads(MANIFESTE.read_text(encoding="utf-8"))


def entete() -> str:
    """Une ligne à reporter dans toute sortie produite sur ce périmètre."""
    m = manifeste()
    return (f"périmètre v{m.get('version', PERIMETRE_VERSION)} — "
            f"{m.get('articles', VOLUME_ATTENDU)} articles — "
            f"figé le {m.get('figé_le', '?')}")


if __name__ == "__main__":
    df = charger()
    print(entete())
    print(f"\n{len(df)} articles")
    print(df["PORTE"].value_counts().to_string())
    print(f"\nRéférences distinctes : {len(references())}")
    print(f"Codes article distincts : {len(codes_article())}")
    print(f"\nColonnes : {list(df.columns)}")
    print(f"\n{df.head(5).to_string(index=False)}")
