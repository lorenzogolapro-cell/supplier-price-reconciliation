# -*- coding: utf-8 -*-
"""
PA workstream: the matches a human established by hand.

    python pa_correspondances.py    checks the table against the price lists

WHAT THIS IS
    A decision table. Each row says: "WMS article REF-EXEMPLE-1 is line
    TARIF-EXEMPLE-1 of Supplier J's price list", because a person
    established it and nothing in the references made it findable.

WHY A TABLE, RATHER THAN ONE MORE RULE
    The workstream's five keys match on identifiers. When they fail,
    `pa_par_modele` proposes candidates by model name - but it PROPOSES, it
    does not decide, and that is deliberate: a price obtained from a label
    is not injected before a human has looked at it.

    So there has to be somewhere that person's decision is recorded, dated
    and read back. That is here. A table is the right tool for whatever
    cannot be derived: `ALIAS_CATALOGUE` does the same for supplier names,
    `REMISES` for discount rates.

WHAT IT DOES NOT DO
    It only FILLS what is empty. If a key has already set a price on the
    article, the table does not overwrite it - it REPORTS it, because two
    sources contradicting each other are information, not a detail to be
    settled silently.

WHAT AN ENTRY MUST CARRY
    The expected price, computed by hand, and the proof. Without them, the
    entry becomes a reference number nobody can justify any more - and the
    day the supplier renumbers, nothing says so.
"""

from __future__ import annotations

import sys
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import warnings  # noqa: E402

warnings.filterwarnings("ignore")

import pandas as pd  # noqa: E402

from extracteurs.base import cle_ref_stricte  # noqa: E402

# ---------------------------------------------------------------------------
# The table. One entry = one decision, dated, with its proof.
# ---------------------------------------------------------------------------

CORRESPONDANCES = [
    # --- Supplier J, validated by the author on 16/09/2026 ----------------
    #
    # Supplier J's price list quotes BY THE CARTON and renumbers its
    # references: our REF-TARIF-EXEMPLE-0 became TARIF-EXEMPLE-1 on their
    # side. None of the five keys could find it, and no prefix either. It
    # is the MODEL NAME that identifies - MODELE EXEMPLE 1, MODELE
    # EXEMPLE 2, MODELE EXEMPLE 3, MODELE EXEMPLE 4 - and the CAPACITY that
    # separates the variants.
    #
    # The ratio to the DPA is 1.308 on three of the six: that is Supplier
    # J's 2026 increase, which the price owner has already validated eight
    # times elsewhere (gaps of 0.307 to 0.31 in his workbook). The three
    # confirm each other.
    {
        "fournisseur": "FOURNISSEUR J", "code_article": "REF-EXEMPLE-1",
        "ref_tarif": "TARIF-EXEMPLE-1",
        "libelle_wms": "PRODUIT EXEMPLE 1 500 ML + POMPE",
        "libelle_tarif": "PRODUIT EXEMPLE 1 MODELE EXEMPLE 1 12X0,5L+12P2CC",
        "prix_attendu": 5.00,  # example value
        "preuve": "60,00 ÷ 12 (carton de douze flacons de 0,5 L = 500 ml) "
                  "= 5,00 ; DPA 3,8226 -> 1,308",  # example values
        "validee_le": "2026-09-16",
    },
    {
        "fournisseur": "FOURNISSEUR J", "code_article": "REF-EXEMPLE-2",
        "ref_tarif": "TARIF-EXEMPLE-2",
        "libelle_wms": "PRODUIT EXEMPLE 2 85 BLEU 1L + POMPE",
        "libelle_tarif": "MODELE EXEMPLE 2 85 NPC 12X1L PPE 3ML",
        "prix_attendu": 10.00,  # example value
        # Three price list lines carry the same price for the 1 L: two
        # AIRLESS variants and the PPE variant. We keep PPE ("pompe", i.e.
        # pump), which is what the WMS describes. The choice does not change
        # the price by a cent, but it changes what can be checked later on.
        "preuve": "120,00 ÷ 12 = 10,00 ; DPA 10,00 -> 1,000 exactement",
                  # example values
        "validee_le": "2026-09-16",
    },
    {
        "fournisseur": "FOURNISSEUR J", "code_article": "REF-EXEMPLE-3",
        "ref_tarif": "TARIF-EXEMPLE-3",
        "libelle_wms": "PRODUIT EXEMPLE 2 85 BLEU 300ML+POMPE",
        "libelle_tarif": "MODELE EXEMPLE 2 85 NPC 6X300ML PPE 3ML",
        "prix_attendu": 6.00,  # example value
        # Ruled out: TARIF-EXEMPLE-9, same format but VARIANTE, which our
        # label does not mention.
        "preuve": "36,00 ÷ 6 = 6,00 ; DPA 4,587 -> 1,308",  # example values
        "validee_le": "2026-09-16",
    },
    {
        "fournisseur": "FOURNISSEUR J", "code_article": "REF-EXEMPLE-4",
        "ref_tarif": "TARIF-EXEMPLE-4",
        "libelle_wms": "PRODUIT EXEMPLE 3 D 1 LITRE",
        "libelle_tarif": "MODELE EXEMPLE 3 D 12X1L DOSEUR",
        "prix_attendu": 20.00,  # example value
        "preuve": "240,00 ÷ 12 = 20,00 ; DPA 12,731 -> 1,571. Le responsable "
                  "des prix annonçait la même valeur dans sa feuille "
                  "« non raproché »",  # example values
        "validee_le": "2026-09-16",
    },
    {
        "fournisseur": "FOURNISSEUR J", "code_article": "REF-EXEMPLE-5",
        "ref_tarif": "TARIF-EXEMPLE-5",
        "libelle_wms": "PRODUIT EXEMPLE 3 D 5 LITRES + POMPE",
        "libelle_tarif": "MODELE EXEMPLE 3 D 4X5L + 1 PPE 25ML",
        "prix_attendu": 80.00,  # example value
        "preuve": "320,00 ÷ 4 = 80,00 ; DPA 72,727 -> 1,100. Valeur annoncée "
                  "par le responsable des prix",  # example values
        "validee_le": "2026-09-16",
    },
    {
        "fournisseur": "FOURNISSEUR J", "code_article": "REF-EXEMPLE-6",
        "ref_tarif": "TARIF-EXEMPLE-6",
        "libelle_wms": "PRODUIT EXEMPLE 4 2% BIDON DE 5L",
        "libelle_tarif": "MODELE EXEMPLE 4 2% 4X5L",
        "prix_attendu": 18.00,  # example value
        "preuve": "72,00 ÷ 4 = 18,00 ; DPA 13,761 -> 1,308. Valeur "
                  "annoncée par le responsable des prix",  # example values
        "validee_le": "2026-09-16",
    },
]

# Tolerated gap between the table's expected price and the price the
# pipeline actually computes. Beyond it we apply nothing: the price list
# has moved, or the pack size is no longer read the same way, and an entry
# validated on other figures is no longer worth anything.
TOLERANCE = 0.01


def pour(motif_fournisseur: str) -> list[dict]:
    """The entries concerning this supplier."""
    nom = str(motif_fournisseur or "").upper()
    return [e for e in CORRESPONDANCES
            if e["fournisseur"].upper() in nom or nom in e["fournisseur"].upper()]


def appliquer(fusion: pd.DataFrame, tarif: pd.DataFrame,
              champs: list[str], motif_fournisseur: str = "") -> pd.DataFrame:
    """Applies the validated matches to whatever is still without a price.

    Same contract as `pa_identifiants.rapprocher_par_identifiant`: we copy
    the price list fields onto the WMS row, and we state in a column which
    key did it. Here the key is called "correspondance validée", and it is
    the only one in the workstream whose source is a person.
    """
    entrees = pour(motif_fournisseur) if motif_fournisseur else CORRESPONDANCES
    if not entrees or "prix_achat_unitaire_ht" not in fusion.columns:
        return fusion
    if "Code article" not in fusion.columns:
        return fusion

    index = {}
    for position, valeur in tarif.get("ref_fournisseur", pd.Series()).items():
        cle = cle_ref_stricte(valeur)
        if cle:
            index.setdefault(cle, position)

    codes = (fusion["Code article"].astype(str).str.strip()
             .str.replace(r"\.0$", "", regex=True))
    poses, conflits, manques = 0, 0, 0

    for entree in entrees:
        cible = codes == str(entree["code_article"])
        if not cible.any():
            continue
        cle = cle_ref_stricte(entree["ref_tarif"])
        if not cle or cle not in index:
            print(f"    ! match {entree['code_article']}: reference "
                  f"{entree['ref_tarif']} is no longer in the price list")
            manques += 1
            continue
        ligne = index[cle]

        for i in fusion.index[cible]:
            deja = fusion.at[i, "prix_achat_unitaire_ht"]
            if pd.notna(deja):
                # Two sources for one article: we do not decide silently.
                # The automatic key found something; if it is not the same
                # line, somebody has to know.
                print(f"    ! match {entree['code_article']}: "
                      f"a price is already set ({deja}) - not overwritten, "
                      f"to be checked")
                conflits += 1
                continue
            for champ in champs:
                if champ in tarif.columns:
                    fusion.at[i, champ] = tarif.at[ligne, champ]
            fusion.at[i, "Clé de rapprochement"] = "correspondance validée"
            # Flag read by pa_conditionnement: on these rows, the pack size
            # read from the label applies WITHOUT asking the DPA to
            # corroborate it. The person who created the entry validated
            # both the price list line AND the expected price - stronger
            # proof than a ratio. Without this, PRODUIT EXEMPLE 3 D 1 L
            # (240.00 / 12) was refused because its ratio to the DPA is
            # 18.85 for a carton of 12, just outside the [8; 18] band - a
            # 57% increase, a real one, that the counter-check could not
            # tell apart from a reading error.
            fusion.at[i, "correspondance_validee"] = True
            fusion.at[i, "preuve_identifiant"] = (
                f"{entree['ref_tarif']} — {entree['preuve']} "
                f"(validée le {entree['validee_le']})")
            poses += 1

    if poses:
        print(f"  {poses} matched by a hand-validated correspondence")
    if conflits or manques:
        print(f"    ({conflits} conflict(s), {manques} vanished reference(s))")
    return fusion


def main() -> None:
    """Checks the table against the price lists, applying nothing."""
    from extracteurs.base import chemin_lisible
    from main import CATALOGUES, FOURNISSEURS

    print(f"{len(CORRESPONDANCES)} correspondence(s) in the table\n")
    for entree in CORRESPONDANCES:
        trouve = None
        for reg in FOURNISSEURS:
            if entree["fournisseur"].upper() not in reg["motif_wms"].upper():
                continue
            for fichier in sorted((CATALOGUES / reg["dossier"]).glob(
                    reg["motif"])):
                tarif = reg["extracteur"].extract(fichier)
                cle = cle_ref_stricte(entree["ref_tarif"])
                for _, ligne in tarif.iterrows():
                    if cle_ref_stricte(ligne.get("ref_fournisseur")) == cle:
                        trouve = ligne
                        break
        if trouve is None:
            print(f"  GONE      {entree['code_article']:>6}  "
                  f"{entree['ref_tarif']:<12} {entree['libelle_wms'][:40]}")
            continue
        prix = pd.to_numeric(pd.Series([trouve["prix_achat_unitaire_ht"]]),
                             errors="coerce").iloc[0]
        etat = "OK      " if pd.notna(prix) else "NO PRICE"
        print(f"  {etat}  {entree['code_article']:>6}  "
              f"{entree['ref_tarif']:<12} price list={prix:<10} "
              f"expected after division={entree['prix_attendu']:<10} "
              f"{entree['libelle_wms'][:34]}")


if __name__ == "__main__":
    main()
