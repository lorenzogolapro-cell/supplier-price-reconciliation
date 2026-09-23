# -*- coding: utf-8 -*-
"""
Tableau de controle de la logique tarifaire, sans rien ecrire sur le disque.

    python controle_calculs.py [nb_lignes]

Affiche pour les premieres lignes du catalogue la chaine complete de calcul
du prix, afin de verifier a la main que :
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

    print("Extraction (aucun fichier ecrit)")
    df = collecter()
    if df.empty:
        print("Aucune ligne extraite.")
        return

    colonnes = [
        "ref_fournisseur", "designation", "tarif_public_ttc", "tva_taux",
        "remise_taux", "prix_achat_unitaire_ht", "prix_recalcule",
        "prix_coherent", "conditionnement", "prix_colis_ht",
        "palier2_qte", "palier2_prix_ht",
    ]

    apercu = df[colonnes].head(nb).copy()
    # Arrondi d'AFFICHAGE uniquement : le DataFrame conserve la precision
    for colonne in ("tarif_public_ttc", "prix_achat_unitaire_ht",
                    "prix_recalcule", "prix_colis_ht", "palier2_prix_ht"):
        apercu[colonne] = apercu[colonne].round(2)
    apercu["designation"] = apercu["designation"].str.slice(0, 34)

    print(f"\n{'=' * 120}")
    print(f"TABLEAU DE CONTROLE - {nb} premieres lignes")
    print("=" * 120)
    print(apercu.to_string(index=False))

    print(f"\n{'=' * 120}")
    print("VERIFICATION DE LA CHAINE DE CALCUL")
    print("=" * 120)
    for _, ligne in df.head(3).iterrows():
        ttc, tva = ligne["tarif_public_ttc"], ligne["tva_taux"]
        remise, net = ligne["remise_taux"], ligne["prix_achat_unitaire_ht"]
        ht = ttc / (1 + tva)
        print(f"  {ligne['ref_fournisseur']:<12} "
              f"{ttc:.2f} TTC / {1 + tva:.3f} = {ht:.2f} HT ; "
              f"x (1 - {remise:.5f}) = {ht * (1 - remise):.2f} "
              f"| fichier : {net:.2f} "
              f"| {'OK' if ligne['prix_coherent'] else 'ECART'}")
        print(f"  {'':<12} colis : {net:.2f} x "
              f"{ligne['conditionnement']:g} = {ligne['prix_colis_ht']:.2f} €")

    print(f"\n{'=' * 120}")
    print("CONTROLE : prix_colis_ht ne doit jamais egaler le prix unitaire "
          "quand le conditionnement > 1")
    print("=" * 120)
    multi = df[df["conditionnement"] > 1]
    faux = multi[
        (multi["prix_colis_ht"] - multi["prix_achat_unitaire_ht"]).abs() < 1e-9
    ]
    print(f"  lignes avec conditionnement > 1 : {len(multi)}")
    print(f"  dont prix colis = prix unitaire : {len(faux)} "
          f"{'(anomalie !)' if len(faux) else '(aucune, correct)'}")

    print(f"\n{'=' * 120}")
    print("FONCTION prix_pour_quantite() - exemples")
    print("=" * 120)
    # Un produit avec paliers et conditionnement > 1
    exemple = df[(df["palier3_prix_ht"].notna()) & (df["conditionnement"] > 1)]
    if not exemple.empty:
        ligne = exemple.iloc[0]
        print(f"  reference {ligne['ref_fournisseur']} - {ligne['designation']}")
        print(f"  prix unitaire {ligne['prix_achat_unitaire_ht']:.2f} € | "
              f"conditionnement {ligne['conditionnement']:g} | "
              f"paliers {ligne['paliers']}")
        for quantite in (1, 3, 10, 50, 200):
            r = prix_pour_quantite(ligne, quantite)
            arrondi = (f" (arrondi depuis {r['quantite_demandee']})"
                       if r["arrondi_conditionnement"] else "")
            print(f"    demande {quantite:>4} -> commande "
                  f"{r['quantite_commandee']:>4.0f}{arrondi:<22} "
                  f"| {r['prix_unitaire_ht']:>7.2f} €/u "
                  f"| total {r['total_ht']:>9.2f} € "
                  f"| {r['palier_applique']}")

    print(f"\n{'=' * 120}")
    print("SYNTHESE QUALITE")
    print("=" * 120)
    stats = calculer_stats(df)
    print(stats.to_string(index=False))

    anomalies = construire_anomalies(df)
    print(f"\n  total lignes d'anomalies : {len(anomalies)}")
    print("\n  repartition par type :")
    print(anomalies["type_anomalie"].value_counts().to_string())

    print("\nAucun fichier ecrit. Lancer `python main.py` pour generer "
          "le fichier consolide.")


if __name__ == "__main__":
    main()
