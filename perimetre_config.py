# -*- coding: utf-8 -*-
"""
Item scope: frozen parameters.

THIS FILE CARRIES BUSINESS DECISIONS, NOT SETTINGS
==================================================
The values below have been validated. They are not parameters to be
tuned, and nobody changes them on their own initiative: not a script,
not an agent, not an "improvement".

A number that gets adjusted to improve a result stops being a
criterion: it becomes an adjustment variable, and the ranking it
produces no longer means anything. If one of them has to change, that
is an explicit decision, it is dated, and the version number below
changes with it.

Every run that reads this file must carry PERIMETRE_VERSION over into
its output: without it, when looking at a ranking, there is no telling
which rules produced it.
"""

from __future__ import annotations

# =====================================================================
# VALIDATED BUSINESS DECISIONS - DO NOT MODIFY
# =====================================================================

# Sales observation window, in months.
# This is the span over which we look at whether an item moved.
FENETRE_VENTE_MOIS = 12

# Minimum number of sales, over the window, for an item to count as
# moving. A single sale is enough.
SEUIL_TOURNE = 1

# How often the ranking is replayed.
CADENCE_RECALCUL = "trimestrielle"

# Rule of entry into the scope (validated on 03/09/2026).
#
#   SCOPE = no exclusion reason
#           AND (sold within 12 months OR purchase received within
#                12 months)
#
# A received purchase is an ENTRY criterion, not merely a confirming
# signal: an item being restocked is a live item, even if it has not
# sold yet. It is the only way to catch new items and those whose sales
# have not come through yet.
#
# THE ORDER MATTERS: the exclusion reasons apply FIRST. An excluded
# item (SE, SV, FO, 19, 79, LO, N903, old nomenclature) does not come
# in, even with a recent purchase.
ACHAT_VAUT_ENTREE = True

# Third door of entry, validated on 08/09/2026: the PRO CATALOGUE.
#
#   The scope keeps what MOVES. The professional catalogue keeps what
#   we PUBLISH. The two do not coincide, and an item published without
#   a purchase price stays an item published without a purchase price,
#   whatever its movement history.
#
# THE ORDER DOES NOT CHANGE: the exclusion reasons still apply FIRST. A
# catalogue reference excluded by a reason (SE, SV, FO, 19, 79, LO,
# N903, old nomenclature) does NOT come in: 58 are in that case, 29 of
# them "discontinued by manufacturer" and carrying that discontinuation
# in their own label. They are documented as "excluded AND in the pro
# catalogue": a discontinued item that stays published is a signal to
# raise with the business, not a row to handle quietly.
CATALOGUE_PRO_VAUT_ENTREE = True

# SCOPE FROZEN. The rule only moves on an explicit decision, and no
# item is ever reclassified on its own initiative.
#
# Expected volume: 5,082 items. Any run that returns another number,
# with the extracts unchanged, is reporting a regression, not an
# evolution.
#   v1.1  4,731  sold OR purchase received
#   v1.2  5,082  + 351 pro catalogue references with no activity
PERIMETRE_FIGE = True
VOLUME_ATTENDU = 5082

# Items with neither a sale nor a purchase leave with the
# SANS_ACTIVITE reason. They STAY in the master data: they are simply
# not worked on. Leaving the scope is not leaving the database.
MOTIF_SANS_ACTIVITE = "SANS_ACTIVITE"

# Reach of each workstream, decided along with the scope:
#   PA  : only work on items inside the scope. Discrepancies outside
#         the scope are documented and closed, not picked up.
#   EAN : EUDAMED runs on the WHOLE master data. The cache is already
#         built, replaying costs nothing, so nothing justifies
#         restricting a free lookup.
CHANTIERS_LIMITES_AU_PERIMETRE = ("PA",)

# EUDAMED is not bounded by the scope: an automated query against an
# already-built cache costs neither money nor human time. Restricting
# to the scope only holds for what costs something: an email to a
# supplier, a count in the warehouse.
EUDAMED_SUR_TOUT_LE_REFERENTIEL = True

# ---------------------------------------------------------------------
# End of the business decisions. What follows derives from them.
# ---------------------------------------------------------------------

# Date these values were validated.
DATE_VALIDATION = "2026-09-03"

# Scope version. To be carried over into any output produced with these
# rules, and to be bumped if one of them changes.
#   1.0  12-month window, threshold of 1 sale, quarterly recompute
#   1.1  a purchase received within 12 months becomes an entry criterion
#   1.2  the pro catalogue becomes a third door of entry
PERIMETRE_VERSION = "1.2"

# The cadence translated into months, to compute the next due date. One
# single source of truth: the cadence stays the wording above, this is
# only its machine-readable form.
CADENCES_EN_MOIS = {
    "mensuelle": 1,
    "trimestrielle": 3,
    "semestrielle": 6,
    "annuelle": 12,
}
CADENCE_RECALCUL_MOIS = CADENCES_EN_MOIS[CADENCE_RECALCUL]


def resume() -> str:
    """The rules in force, on one line, for an output header."""
    entree = "sale OR purchase received" if ACHAT_VAUT_ENTREE else "sale"
    if CATALOGUE_PRO_VAUT_ENTREE:
        entree += " OR pro catalogue"
    return (f"scope v{PERIMETRE_VERSION} - "
            f"{FENETRE_VENTE_MOIS}-month window, "
            f"threshold {SEUIL_TOURNE} sale, "
            f"entry: {entree}, "
            f"{CADENCE_RECALCUL} recompute - "
            f"validated on {DATE_VALIDATION}")


if __name__ == "__main__":
    print(resume())
