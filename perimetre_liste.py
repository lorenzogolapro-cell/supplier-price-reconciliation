# -*- coding: utf-8 -*-
"""
The frozen scope, read-only.

This is the entry point for EVERY session working on the scope: the
purchase-price workstream, the EAN workstream, or anything else.

    from perimetre_liste import charger, references

    liste = charger()            # DataFrame, 4,731 rows
    refs = references()          # set of internal references
    if ref in refs: ...

Why NOT to recompute
--------------------
The scope is frozen. Recomputing it at every session means risking that
two sessions work on two different populations without anyone noticing:
all it takes is one extract having been regenerated in between.

This module reads the list written once by `figer_perimetre.py`. It
recomputes nothing, ever. If the list is missing, it says so and stops
rather than rebuilding it on the fly.

If the scope has to change, that is a decision: bump PERIMETRE_VERSION
and rerun `figer_perimetre.py --refaire`.
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
    f"The frozen scope cannot be found: {LISTE}\n"
    f"Produce it with:  python figer_perimetre.py\n"
    f"Do NOT recompute it inside your own script: two sessions that "
    f"both recompute can diverge."
)


@lru_cache(maxsize=1)
def charger() -> pd.DataFrame:
    """The frozen scope. Read once, cached for the session."""
    if not LISTE.exists():
        raise SystemExit(_ABSENTE)
    df = pd.read_csv(LISTE, dtype=str)
    for colonne in ("VENTE_12M", "ACHAT_12M"):
        if colonne in df.columns:
            df[colonne] = df[colonne].astype(str).str.lower() == "true"
    if len(df) != VOLUME_ATTENDU:
        raise SystemExit(
            f"The frozen scope holds {len(df)} items, "
            f"{VOLUME_ATTENDU} expected. The list has been modified "
            f"outside figer_perimetre.py: check before going any further.")
    return df


@lru_cache(maxsize=1)
def references() -> frozenset:
    """The internal references of the scope, for a membership test."""
    return frozenset(charger()["Réf. interne"].dropna())


@lru_cache(maxsize=1)
def codes_article() -> frozenset:
    """The item codes of the scope.

    To be used ONLY to match one WMS export against another. Towards any
    outside source, the key remains the internal reference.
    """
    return frozenset(charger()["Code article"].dropna())


def manifeste() -> dict:
    """Version, freeze date and fingerprint of the sources."""
    if not MANIFESTE.exists():
        return {}
    return json.loads(MANIFESTE.read_text(encoding="utf-8"))


def entete() -> str:
    """One line to carry over into any output produced on this scope."""
    m = manifeste()
    return (f"scope v{m.get('version', PERIMETRE_VERSION)} - "
            f"{m.get('articles', VOLUME_ATTENDU)} items - "
            f"frozen on {m.get('figé_le', '?')}")


if __name__ == "__main__":
    df = charger()
    print(entete())
    print(f"\n{len(df)} items")
    print(df["PORTE"].value_counts().to_string())
    print(f"\nDistinct internal references: {len(references())}")
    print(f"Distinct item codes: {len(codes_article())}")
    print(f"\nColumns: {list(df.columns)}")
    print(f"\n{df.head(5).to_string(index=False)}")
