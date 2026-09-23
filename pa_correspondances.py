# -*- coding: utf-8 -*-
"""
Chantier PA — les correspondances qu'un humain a posées à la main.

    python pa_correspondances.py        contrôle la table contre les tarifs

CE QUE C'EST
    Une table de décisions. Chaque ligne dit : « l'article REF-EXEMPLE-1 du
    WMS est la ligne TARIF-EXEMPLE-1 du tarif du Fournisseur J », parce
    qu'une personne l'a établi et que rien dans les références ne permettait
    de le trouver.

POURQUOI UNE TABLE, ET PAS UNE RÈGLE DE PLUS
    Les cinq clés du chantier rapprochent par identifiant. Quand elles
    échouent, `pa_par_modele` propose des candidats par nom de modèle — mais
    il PROPOSE, il ne tranche pas, et c'est délibéré : un prix obtenu par
    libellé ne s'injecte pas sans qu'un humain l'ait regardé.

    Il faut donc un endroit où la décision de cet humain se dépose, se date,
    et se relit. C'est ici. Une table est le bon outil pour ce qui ne se
    déduit pas : `ALIAS_CATALOGUE` fait la même chose pour les noms de
    fournisseurs, `REMISES` pour les taux.

CE QU'ELLE NE FAIT PAS
    Elle ne COMBLE que ce qui est vide. Si une clé a déjà posé un prix sur
    l'article, la table ne l'écrase pas — elle le SIGNALE, parce que deux
    sources qui se contredisent sont une information, pas un détail à
    trancher en silence.

CE QU'UNE ENTRÉE DOIT PORTER
    Le prix attendu, calculé à la main, et la preuve. Sans eux, l'entrée
    devient un numéro de référence que plus personne ne sait justifier —
    et le jour où le fournisseur renumérote, rien ne le dit.
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
# La table. Une entrée = une décision, datée, avec sa preuve.
# ---------------------------------------------------------------------------

CORRESPONDANCES = [
    # --- Fournisseur J, validé par l'auteur le 16/09/2026 -----------------
    #
    # Le tarif du Fournisseur J cote AU CARTON et renumérote ses références :
    # notre REF-TARIF-EXEMPLE-0 est devenue TARIF-EXEMPLE-1 chez eux. Aucune
    # des cinq clés ne pouvait le trouver, et aucun préfixe non plus. C'est
    # le NOM DE MODÈLE qui identifie — MODELE EXEMPLE 1, MODELE EXEMPLE 2,
    # MODELE EXEMPLE 3, MODELE EXEMPLE 4 — et la CONTENANCE qui sépare les
    # déclinaisons.
    #
    # Le rapport au DPA vaut 1,308 sur trois des six : c'est la hausse 2026
    # du Fournisseur J, que le responsable des prix a déjà validée huit fois
    # ailleurs (écarts 0,307 à 0,31 dans son classeur). Les trois se
    # confirment mutuellement.
    {
        "fournisseur": "FOURNISSEUR J", "code_article": "REF-EXEMPLE-1",
        "ref_tarif": "TARIF-EXEMPLE-1",
        "libelle_wms": "PRODUIT EXEMPLE 1 500 ML + POMPE",
        "libelle_tarif": "PRODUIT EXEMPLE 1 MODELE EXEMPLE 1 12X0,5L+12P2CC",
        "prix_attendu": 5.00,  # valeur d'exemple
        "preuve": "60,00 ÷ 12 (carton de douze flacons de 0,5 L = 500 ml) "
                  "= 5,00 ; DPA 3,8226 -> 1,308",  # valeurs d'exemple
        "validee_le": "2026-09-16",
    },
    {
        "fournisseur": "FOURNISSEUR J", "code_article": "REF-EXEMPLE-2",
        "ref_tarif": "TARIF-EXEMPLE-2",
        "libelle_wms": "PRODUIT EXEMPLE 2 85 BLEU 1L + POMPE",
        "libelle_tarif": "MODELE EXEMPLE 2 85 NPC 12X1L PPE 3ML",
        "prix_attendu": 10.00,  # valeur d'exemple
        # Trois lignes du tarif portent le même prix pour le 1 L : deux
        # variantes AIRLESS et la variante PPE. On retient PPE — « pompe » —
        # qui est ce que le WMS décrit. Le choix ne change pas le prix d'un
        # centime, mais il change ce qu'on pourra vérifier plus tard.
        "preuve": "120,00 ÷ 12 = 10,00 ; DPA 10,00 -> 1,000 exactement",
                  # valeurs d'exemple
        "validee_le": "2026-09-16",
    },
    {
        "fournisseur": "FOURNISSEUR J", "code_article": "REF-EXEMPLE-3",
        "ref_tarif": "TARIF-EXEMPLE-3",
        "libelle_wms": "PRODUIT EXEMPLE 2 85 BLEU 300ML+POMPE",
        "libelle_tarif": "MODELE EXEMPLE 2 85 NPC 6X300ML PPE 3ML",
        "prix_attendu": 6.00,  # valeur d'exemple
        # Écarté : TARIF-EXEMPLE-9, même format mais VARIANTE, que notre
        # libellé ne mentionne pas.
        "preuve": "36,00 ÷ 6 = 6,00 ; DPA 4,587 -> 1,308",  # valeurs d'exemple
        "validee_le": "2026-09-16",
    },
    {
        "fournisseur": "FOURNISSEUR J", "code_article": "REF-EXEMPLE-4",
        "ref_tarif": "TARIF-EXEMPLE-4",
        "libelle_wms": "PRODUIT EXEMPLE 3 D 1 LITRE",
        "libelle_tarif": "MODELE EXEMPLE 3 D 12X1L DOSEUR",
        "prix_attendu": 20.00,  # valeur d'exemple
        "preuve": "240,00 ÷ 12 = 20,00 ; DPA 12,731 -> 1,571. Le responsable "
                  "des prix annonçait la même valeur dans sa feuille "
                  "« non raproché »",  # valeurs d'exemple
        "validee_le": "2026-09-16",
    },
    {
        "fournisseur": "FOURNISSEUR J", "code_article": "REF-EXEMPLE-5",
        "ref_tarif": "TARIF-EXEMPLE-5",
        "libelle_wms": "PRODUIT EXEMPLE 3 D 5 LITRES + POMPE",
        "libelle_tarif": "MODELE EXEMPLE 3 D 4X5L + 1 PPE 25ML",
        "prix_attendu": 80.00,  # valeur d'exemple
        "preuve": "320,00 ÷ 4 = 80,00 ; DPA 72,727 -> 1,100. Valeur annoncée "
                  "par le responsable des prix",  # valeurs d'exemple
        "validee_le": "2026-09-16",
    },
    {
        "fournisseur": "FOURNISSEUR J", "code_article": "REF-EXEMPLE-6",
        "ref_tarif": "TARIF-EXEMPLE-6",
        "libelle_wms": "PRODUIT EXEMPLE 4 2% BIDON DE 5L",
        "libelle_tarif": "MODELE EXEMPLE 4 2% 4X5L",
        "prix_attendu": 18.00,  # valeur d'exemple
        "preuve": "72,00 ÷ 4 = 18,00 ; DPA 13,761 -> 1,308. Valeur "
                  "annoncée par le responsable des prix",  # valeurs d'exemple
        "validee_le": "2026-09-16",
    },
]

# Écart toléré entre le prix attendu de la table et le prix réellement
# calculé par le pipeline. Au-delà, on n'applique rien : le tarif a bougé,
# ou le conditionnement n'est plus lu pareil, et une entrée validée sur
# d'autres chiffres ne vaut plus.
TOLERANCE = 0.01


def pour(motif_fournisseur: str) -> list[dict]:
    """Les entrées qui concernent ce fournisseur."""
    nom = str(motif_fournisseur or "").upper()
    return [e for e in CORRESPONDANCES
            if e["fournisseur"].upper() in nom or nom in e["fournisseur"].upper()]


def appliquer(fusion: pd.DataFrame, tarif: pd.DataFrame,
              champs: list[str], motif_fournisseur: str = "") -> pd.DataFrame:
    """Pose les correspondances validées sur ce qui reste sans prix.

    Même contrat que `pa_identifiants.rapprocher_par_identifiant` : on copie
    les champs du tarif sur la ligne du WMS, et on dit en colonne par quelle
    clé. Ici la clé s'appelle « correspondance validée », et c'est la seule
    du chantier dont la source est une personne.
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
            print(f"    ! correspondance {entree['code_article']} : la "
                  f"référence {entree['ref_tarif']} n'est plus au tarif")
            manques += 1
            continue
        ligne = index[cle]

        for i in fusion.index[cible]:
            deja = fusion.at[i, "prix_achat_unitaire_ht"]
            if pd.notna(deja):
                # Deux sources pour un meme article : on ne tranche pas en
                # silence. La cle automatique a trouve quelque chose ; si ce
                # n'est pas la meme ligne, quelqu'un doit le savoir.
                print(f"    ! correspondance {entree['code_article']} : "
                      f"un prix est déjà posé ({deja}) — non écrasé, "
                      f"à vérifier")
                conflits += 1
                continue
            for champ in champs:
                if champ in tarif.columns:
                    fusion.at[i, champ] = tarif.at[ligne, champ]
            fusion.at[i, "Clé de rapprochement"] = "correspondance validée"
            # Marque lue par pa_conditionnement : sur ces lignes, le
            # conditionnement lu au libellé s'applique SANS demander au DPA
            # de le corroborer. La personne qui a posé l'entrée a validé la
            # ligne du tarif ET le prix attendu — c'est une preuve plus
            # forte qu'un rapport. Sans cela, PRODUIT EXEMPLE 3 D 1 L
            # (240,00 ÷ 12) était refusé parce que son rapport au DPA vaut
            # 18,85 pour un carton de 12, soit juste hors de la fourchette
            # [8 ; 18] — une hausse de 57 %, réelle, que la contre-épreuve
            # ne pouvait pas distinguer d'une erreur de lecture.
            fusion.at[i, "correspondance_validee"] = True
            fusion.at[i, "preuve_identifiant"] = (
                f"{entree['ref_tarif']} — {entree['preuve']} "
                f"(validée le {entree['validee_le']})")
            poses += 1

    if poses:
        print(f"  {poses} rapprochés par correspondance validée à la main")
    if conflits or manques:
        print(f"    ({conflits} conflit(s), {manques} référence(s) disparue(s))")
    return fusion


def main() -> None:
    """Contrôle la table contre les tarifs, sans rien appliquer."""
    from extracteurs.base import chemin_lisible
    from main import CATALOGUES, FOURNISSEURS

    print(f"{len(CORRESPONDANCES)} correspondance(s) dans la table\n")
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
            print(f"  DISPARUE  {entree['code_article']:>6}  "
                  f"{entree['ref_tarif']:<12} {entree['libelle_wms'][:40]}")
            continue
        prix = pd.to_numeric(pd.Series([trouve["prix_achat_unitaire_ht"]]),
                             errors="coerce").iloc[0]
        etat = "OK      " if pd.notna(prix) else "SANS PRIX"
        print(f"  {etat}  {entree['code_article']:>6}  "
              f"{entree['ref_tarif']:<12} tarif={prix:<10} "
              f"attendu après division={entree['prix_attendu']:<10} "
              f"{entree['libelle_wms'][:34]}")


if __name__ == "__main__":
    main()
