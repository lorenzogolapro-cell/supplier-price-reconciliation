# -*- coding: utf-8 -*-
"""
Controles qualite sur le catalogue consolide.

Produit trois choses :
  - une synthese lisible (indicateur / valeur / pourcentage)
  - une table unique d'anomalies typees, triee par gravite
  - le detail par categorie, pour le rapport complet
"""

from __future__ import annotations

import pandas as pd

from extracteurs.base import anomalies_paliers, diagnostic_ean, ean_est_valide

# Gravite des anomalies : plus le chiffre est bas, plus c'est urgent.
# Un prix nul empeche toute commande, un doublon d'EAN fausse les
# identifications produit : ce sont les deux cas bloquants.
GRAVITE = {
    "Prix d'achat nul ou negatif": 1,
    "Doublon EAN (produits differents)": 2,
    "Doublon EAN": 3,
    "Prix incoherent": 4,
    "Palier aberrant": 5,
    "Doublon reference": 6,
    "EAN invalide": 7,
    "EAN manquant": 8,
    "Conditionnement suppose": 9,
}

COLONNES_ANOMALIE = [
    "type_anomalie",
    "ref_fournisseur",
    "designation",
    "ean",
    "prix_achat_unitaire_ht",
    "conditionnement",
    "detail",
    "fournisseur",
]


def _ean_presents(df: pd.DataFrame) -> pd.Series:
    return df["ean"].notna() & (df["ean"].astype(str) != "")


# ---------------------------------------------------------------------------
# Table unique des anomalies
# ---------------------------------------------------------------------------

def construire_anomalies(df: pd.DataFrame) -> pd.DataFrame:
    """Une ligne par anomalie constatee, avec son type en clair.

    Un meme produit peut apparaitre plusieurs fois s'il cumule plusieurs
    problemes : c'est voulu, chaque ligne correspond a une action a mener.
    """
    lignes = []

    def ajouter(masque: pd.Series, type_anomalie: str, detail=None):
        """Empile les lignes correspondant a un masque booleen."""
        if not masque.any():
            return
        extrait = df.loc[masque].copy()
        extrait["type_anomalie"] = type_anomalie
        if detail is None:
            extrait["detail"] = None
        elif isinstance(detail, str):
            extrait["detail"] = detail
        else:
            extrait["detail"] = detail  # Series alignee sur l'index
        lignes.append(extrait)

    presents = _ean_presents(df)

    # --- prix -------------------------------------------------------------
    prix = df["prix_achat_unitaire_ht"]
    ajouter(prix.isna(), "Prix d'achat nul ou negatif", "prix absent du tarif")
    ajouter(prix.notna() & (prix <= 0), "Prix d'achat nul ou negatif",
            "prix inferieur ou egal a zero")

    incoherent = df["prix_coherent"] == False  # noqa: E712 (None doit rester exclu)
    if incoherent.any():
        detail = df.loc[incoherent].apply(
            lambda r: f"recalcule {r['prix_recalcule']:.2f} € "
                      f"contre {r['prix_achat_unitaire_ht']:.2f} € "
                      f"(ecart {r['ecart_prix']:.2f} €)",
            axis=1,
        )
        ajouter(incoherent, "Prix incoherent", detail)

    # --- paliers ----------------------------------------------------------
    def controler(ligne):
        return anomalies_paliers(
            ligne["prix_achat_unitaire_ht"],
            (ligne["palier2_qte"], ligne["palier2_prix_ht"]),
            (ligne["palier3_qte"], ligne["palier3_prix_ht"]),
        )

    problemes = df.apply(controler, axis=1)
    masque_palier = problemes.map(bool)
    if masque_palier.any():
        ajouter(masque_palier, "Palier aberrant",
                problemes.loc[masque_palier].map(" ; ".join))

    # --- conditionnement --------------------------------------------------
    ajouter(df["conditionnement_suppose"] == True,  # noqa: E712
            "Conditionnement suppose", "absent du tarif, ramene a 1")

    # --- EAN --------------------------------------------------------------
    ajouter(~presents, "EAN manquant")

    invalides = presents & ~df["ean"].map(
        lambda e: ean_est_valide(e) if isinstance(e, str) else False
    )
    if invalides.any():
        ajouter(invalides, "EAN invalide", df.loc[invalides, "ean"].map(diagnostic_ean))

    # --- doublons ---------------------------------------------------------
    refs = df["ref_fournisseur"].duplicated(keep=False)
    if refs.any():
        groupes = df.loc[refs].groupby("ref_fournisseur")
        divergent = (
            groupes["prix_achat_unitaire_ht"].transform("nunique") > 1
        ).reindex(df.index, fill_value=False)
        ajouter(refs & divergent, "Doublon reference",
                "meme reference, prix differents")
        ajouter(refs & ~divergent, "Doublon reference",
                "reference repetee, valeurs identiques")

    sous_ean = df.loc[presents]
    doubles = sous_ean["ean"].duplicated(keep=False)
    if doubles.any():
        nb_designations = sous_ean.loc[doubles].groupby("ean")[
            "designation"
        ].transform("nunique")
        produits_differents = (nb_designations > 1).reindex(
            df.index, fill_value=False
        )
        masque = presents & doubles.reindex(df.index, fill_value=False)
        ajouter(masque & produits_differents,
                "Doublon EAN (produits differents)",
                "meme code-barres pour des produits distincts")
        ajouter(masque & ~produits_differents, "Doublon EAN",
                "meme code-barres, meme designation")

    if not lignes:
        return pd.DataFrame(columns=COLONNES_ANOMALIE)

    anomalies = pd.concat(lignes, ignore_index=True)
    anomalies["_gravite"] = anomalies["type_anomalie"].map(GRAVITE).fillna(99)
    anomalies = anomalies.sort_values(
        ["_gravite", "ref_fournisseur"]
    ).drop(columns="_gravite")

    return anomalies[[c for c in COLONNES_ANOMALIE if c in anomalies.columns]]


# ---------------------------------------------------------------------------
# Synthese
# ---------------------------------------------------------------------------

def calculer_stats(df: pd.DataFrame) -> pd.DataFrame:
    """Synthese qualite en tableau lisible : indicateur / valeur / %."""
    total = len(df)
    presents = _ean_presents(df)
    prix = df["prix_achat_unitaire_ht"]

    valides = df.loc[presents, "ean"].map(ean_est_valide)
    nb_invalides = int((~valides).sum()) if presents.any() else 0

    problemes_paliers = df.apply(
        lambda r: bool(
            anomalies_paliers(
                r["prix_achat_unitaire_ht"],
                (r["palier2_qte"], r["palier2_prix_ht"]),
                (r["palier3_qte"], r["palier3_prix_ht"]),
            )
        ),
        axis=1,
    )

    avec_palier = df["palier2_prix_ht"].notna() | df["palier3_prix_ht"].notna()

    indicateurs = [
        ("Lignes extraites", total),
        ("EAN presents", int(presents.sum())),
        ("EAN manquants", int((~presents).sum())),
        ("EAN invalides (cle de controle)", nb_invalides),
        ("Prix d'achat manquants", int(prix.isna().sum())),
        ("Prix d'achat nuls ou negatifs", int((prix.notna() & (prix <= 0)).sum())),
        ("Prix incoherents (ecart > 0,05 €)",
         int((df["prix_coherent"] == False).sum())),  # noqa: E712
        ("Coherence non verifiable (donnee manquante)",
         int(df["prix_coherent"].isna().sum())),
        ("Conditionnements supposes a 1",
         int((df["conditionnement_suppose"] == True).sum())),  # noqa: E712
        ("Lignes avec au moins un palier degressif", int(avec_palier.sum())),
        ("Paliers aberrants", int(problemes_paliers.sum())),
        ("Lignes en doublon de reference",
         int(df["ref_fournisseur"].duplicated(keep=False).sum())),
        ("Lignes en doublon d'EAN",
         int(df.loc[presents, "ean"].duplicated(keep=False).sum())),
    ]

    lignes = [
        {
            "indicateur": nom,
            "valeur": valeur,
            "pourcentage": round(100 * valeur / total, 2) if total else 0.0,
        }
        for nom, valeur in indicateurs
    ]

    # Repartition des taux de TVA rencontres
    for taux, nombre in df["tva_taux"].value_counts(dropna=False).items():
        libelle = "non renseigne" if pd.isna(taux) else f"{taux:.1%}"
        lignes.append(
            {
                "indicateur": f"TVA {libelle}",
                "valeur": int(nombre),
                "pourcentage": round(100 * nombre / total, 2) if total else 0.0,
            }
        )

    return pd.DataFrame(lignes)


# ---------------------------------------------------------------------------
# Detail par categorie, pour le rapport complet
# ---------------------------------------------------------------------------

def detail_anomalies(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """Detail ligne a ligne, un onglet par famille d'anomalie."""
    colonnes = [
        "fournisseur", "ref_fournisseur", "designation", "ean",
        "prix_achat_unitaire_ht", "conditionnement", "fichier_source",
    ]
    presents = _ean_presents(df)

    invalides = presents & ~df["ean"].map(
        lambda e: ean_est_valide(e) if isinstance(e, str) else False
    )
    ean_invalides = df.loc[invalides, colonnes].copy()
    if not ean_invalides.empty:
        ean_invalides["longueur_ean"] = ean_invalides["ean"].str.len()
        ean_invalides["diagnostic"] = ean_invalides["ean"].map(diagnostic_ean)
        ean_invalides = ean_invalides.sort_values(["diagnostic", "ref_fournisseur"])

    incoherences = df.loc[
        df["prix_coherent"] == False,  # noqa: E712
        ["fournisseur", "ref_fournisseur", "designation", "tarif_public_ttc",
         "tva_taux", "remise_taux", "prix_achat_unitaire_ht", "prix_recalcule",
         "ecart_prix"],
    ].sort_values("ecart_prix", ascending=False)

    refs = df.loc[df["ref_fournisseur"].duplicated(keep=False), colonnes].copy()
    if not refs.empty:
        groupes = refs.groupby(["fournisseur", "ref_fournisseur"])
        refs["prix_divergent"] = (
            groupes["prix_achat_unitaire_ht"].transform("nunique") > 1
        ).map({True: "OUI", False: "non"})
        refs = refs.sort_values(
            ["prix_divergent", "ref_fournisseur"], ascending=[False, True]
        )

    sous_ean = df.loc[presents]
    doubles = sous_ean.loc[sous_ean["ean"].duplicated(keep=False), colonnes].copy()
    if not doubles.empty:
        nb = doubles.groupby("ean")["designation"].transform("nunique")
        doubles["produits_differents"] = (nb > 1).map({True: "OUI", False: "non"})
        doubles = doubles.sort_values(
            ["produits_differents", "ean"], ascending=[False, True]
        )

    return {
        "EAN invalides": ean_invalides,
        "EAN manquants": df.loc[~presents, colonnes],
        "Incoherences prix": incoherences,
        "Doublons reference": refs,
        "Doublons EAN": doubles,
    }
