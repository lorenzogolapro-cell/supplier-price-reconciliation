# -*- coding: utf-8 -*-
"""
When a code coming from EUDAMED is worth keeping.

EUDAMED is not at fault: the way we queried it is. We submitted the
article's reference, and the registry returns the device carrying that
reference, at ANY manufacturer in Europe. A four-digit reference
designates hundreds of them.

What that gives in the file, with nothing made up:

    CHAUSSURE MODELE-B 39  -> ref. "2975" -> 5012345678900
    CHAUSSURE MODELE-B 40  -> ref. "2975" -> 5012345678900
    CHAUSSURE MODELE-B 41  -> ref. "2975" -> 5012345678900

Three shoe sizes under a single code, and a British prefix on a French
shoe.

The measurement
---------------
Every code confronted with the GS1 prefix found in the price list of the
SAME supplier (V16, 557 testable codes):

    reference of 3 characters    28 codes     0 % agreement
                 4               23           4 %
                 5               27           7 %
                 6              148          11 %
                 7              111          30 %
                 8              114           8 %
                 9 and above     96          96 %

The break is sharp and owes nothing to chance: beyond nine characters a
reference is distinctive, it brings back a single device and it is the
right one. Below that, it is a lottery.

The test has a bias worth knowing about: EUDAMED indexes the
MANUFACTURER, and a supplier that distributes without manufacturing will
legitimately carry a foreign prefix. That bias works against EUDAMED
more than against the other sources, but it explains neither the 0 % at
three characters nor the jump to 96 % as soon as nine is passed.

What the rule does NOT do
-------------------------
It deletes nothing. The discarded codes go to `perimetre/retours/`, from
where they can be picked up again the day we know how to check them. It
does not touch "EAN distributeur - Vn.xlsx" either, which stays the log
of what the pipeline produced: the sorting happens on the way OUT, in
what gets read and what gets taken to the warehouse.
"""

from __future__ import annotations

import pandas as pd

# Nine, because that is where the measurement tips over, not because it
# is a round number. Relaxing it to 8 would let in 114 codes at 8 %
# agreement.
LONGUEUR_MINIMALE = 9

SOURCE = "eudamed"

MOTIF = (f"référence de moins de {LONGUEUR_MINIMALE} caractères : EUDAMED "
         f"a rendu le dispositif d'un autre fabricant")


def reference_fiable(reference) -> bool:
    """Was this reference distinctive enough to query EUDAMED with?"""
    if reference is None or (isinstance(reference, float) and pd.isna(reference)):
        return False
    return len(str(reference).strip()) >= LONGUEUR_MINIMALE


def douteux(df: pd.DataFrame, colonne_source: str = "Source",
            colonne_reference: str = "Réf. fournisseur") -> pd.Series:
    """The rows whose code comes from EUDAMED on too short a reference.

    Returns a mask, never a truncated frame: it is up to the caller to
    decide whether to discard, to downgrade, or simply to count.
    """
    if colonne_source not in df.columns:
        return pd.Series(False, index=df.index)
    vient_deudamed = df[colonne_source].fillna("").str.strip().eq(SOURCE)
    trop_court = ~df[colonne_reference].map(reference_fiable)
    return vient_deudamed & trop_court
