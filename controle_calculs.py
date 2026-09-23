# -*- coding: utf-8 -*-
"""
Control board for the pricing logic, writing nothing to disk.

    python controle_calculs.py [nb_lignes]

Prints, for the first rows of the catalogue, the full price computation
chain, so that it can be checked by hand that:
    prix_recalcule = (tarif_public_ttc / (1 + tva_taux)) x (1 - remise_taux)
    prix_colis_ht  = prix_achat_unitaire_ht x conditionnement
"""

from __future__ import annotations

import sys
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import pandas as pd  # noqa: E402

from extracteurs.base import prix_pour_quantite  # noqa: E402
from main import collecter  # noqa: E402
from qualite import calculer_stats, construire_anomalies  # noqa: E402

pd.set_option("display.width", 260)
pd.set_option("display.max_columns", None)


def main() -> None:
    nb = int(sys.argv[1]) if len(sys.argv) > 1 else 15

    print("Extraction (no file written)")
    df = collecter()
    if df.empty:
        print("No row extracted.")
        return

    colonnes = [
        "ref_fournisseur", "designation", "tarif_public_ttc", "tva_taux",
        "remise_taux", "prix_achat_unitaire_ht", "prix_recalcule",
        "prix_coherent", "conditionnement", "prix_colis_ht",
        "palier2_qte", "palier2_prix_ht",
    ]

    apercu = df[colonnes].head(nb).copy()
    # DISPLAY rounding only: the DataFrame keeps its full precision
    for colonne in ("tarif_public_ttc", "prix_achat_unitaire_ht",
                    "prix_recalcule", "prix_colis_ht", "palier2_prix_ht"):
        apercu[colonne] = apercu[colonne].round(2)
    apercu["designation"] = apercu["designation"].str.slice(0, 34)

    print(f"\n{'=' * 120}")
    print(f"CONTROL BOARD - first {nb} rows")
    print("=" * 120)
    print(apercu.to_string(index=False))

    print(f"\n{'=' * 120}")
    print("CHECKING THE COMPUTATION CHAIN")
    print("=" * 120)
    for _, ligne in df.head(3).iterrows():
        ttc, tva = ligne["tarif_public_ttc"], ligne["tva_taux"]
        remise, net = ligne["remise_taux"], ligne["prix_achat_unitaire_ht"]
        ht = ttc / (1 + tva)
        print(f"  {ligne['ref_fournisseur']:<12} "
              f"{ttc:.2f} incl. VAT / {1 + tva:.3f} = {ht:.2f} excl. VAT ; "
              f"x (1 - {remise:.5f}) = {ht * (1 - remise):.2f} "
              f"| file: {net:.2f} "
              f"| {'OK' if ligne['prix_coherent'] else 'GAP'}")
        print(f"  {'':<12} pack: {net:.2f} x "
              f"{ligne['conditionnement']:g} = {ligne['prix_colis_ht']:.2f} €")

    print(f"\n{'=' * 120}")
    print("CHECK: prix_colis_ht must never equal the unit price "
          "when the pack size is > 1")
    print("=" * 120)
    multi = df[df["conditionnement"] > 1]
    faux = multi[
        (multi["prix_colis_ht"] - multi["prix_achat_unitaire_ht"]).abs() < 1e-9
    ]
    print(f"  rows with a pack size > 1: {len(multi)}")
    print(f"  of which pack price = unit price: {len(faux)} "
          f"{'(anomaly!)' if len(faux) else '(none, correct)'}")

    print(f"\n{'=' * 120}")
    print("prix_pour_quantite() FUNCTION - examples")
    print("=" * 120)
    # A product with price breaks and a pack size > 1
    exemple = df[(df["palier3_prix_ht"].notna()) & (df["conditionnement"] > 1)]
    if not exemple.empty:
        ligne = exemple.iloc[0]
        print(f"  reference {ligne['ref_fournisseur']} - {ligne['designation']}")
        print(f"  unit price {ligne['prix_achat_unitaire_ht']:.2f} € | "
              f"pack size {ligne['conditionnement']:g} | "
              f"price breaks {ligne['paliers']}")
        for quantite in (1, 3, 10, 50, 200):
            r = prix_pour_quantite(ligne, quantite)
            arrondi = (f" (rounded up from {r['quantite_demandee']})"
                       if r["arrondi_conditionnement"] else "")
            print(f"    asked {quantite:>4} -> ordered "
                  f"{r['quantite_commandee']:>4.0f}{arrondi:<22} "
                  f"| {r['prix_unitaire_ht']:>7.2f} €/u "
                  f"| total {r['total_ht']:>9.2f} € "
                  f"| {r['palier_applique']}")

    print(f"\n{'=' * 120}")
    print("QUALITY SUMMARY")
    print("=" * 120)
    stats = calculer_stats(df)
    print(stats.to_string(index=False))

    anomalies = construire_anomalies(df)
    print(f"\n  total anomaly rows: {len(anomalies)}")
    print("\n  breakdown by type:")
    print(anomalies["type_anomalie"].value_counts().to_string())

    print("\nNo file written. Run `python main.py` to generate "
          "the consolidated file.")


if __name__ == "__main__":
    main()
