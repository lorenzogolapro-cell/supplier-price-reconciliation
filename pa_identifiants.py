# -*- coding: utf-8 -*-
"""
Chantier PA — réparer les identifiants abîmés avant tout rapprochement flou.

    python pa_identifiants.py               diagnostic sur le périmètre
    python pa_identifiants.py FOURNISSEUR_E un fournisseur

POURQUOI AVANT LE LIBELLÉ
    Une correspondance démontrée par la clé de contrôle d'un EAN est une
    correspondance EXACTE, pas un candidat : la somme pondérée 3/1 de GS1 ne
    tombe juste par hasard qu'une fois sur dix. Elle s'applique donc, là où un
    rapprochement par libellé ne peut que se proposer.

CE QU'ON FAIT
    1. Plusieurs identifiants dans une seule cellule.
       « 1234567890128+1234567890135 » : deux EAN valides séparés par un « + ».
       On les découpe et on essaie chacun. Aucun chiffre n'est ajouté.

    2. GTIN-14 dont l'indicateur de tête vaut zéro.
       Les treize chiffres de l'EAN y sont déjà tous ; on retire le bourrage.

ON NE RECONSTRUIT AUCUN IDENTIFIANT (décision du 16/09)
    Trois réparations vivaient ici — EAN tronqué dont on recalculait la clé,
    zéro de tête restitué, clé fausse corrigée. Toutes retirées. Une clé de
    contrôle qui tombe juste démontre moins qu'il n'y paraît : elle tombe
    juste une fois sur dix au hasard, et sur des milliers de références cela
    arrive assez pour poser des prix faux et muets. Un identifiant abîmé se
    SIGNALE ; il repart vers le rapprochement par libellé, qui propose au
    lieu d'appliquer.

CE QU'ON NE RÉPARE PAS NON PLUS
    Le socle de référence — le radical numérique d'au moins six chiffres, qui
    rapproche, chez le Fournisseur J, « 1234567 » de « 1234567HR ». Chez le
    responsable des prix il est SIGNALÉ,
    jamais appliqué, et on garde sa règle : un suffixe de gamme distingue
    souvent deux produits réels. Il ne devient exploitable que corroboré par le
    prix, et il ressort alors en « À RELIRE », jamais en injectable.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

from extracteurs.base import (  # noqa: E402
    cle_ref_stricte,
    ean_est_valide,
)

# Un « + » sépare deux identifiants ; deux espaces ou plus aussi, quand la
# saisie a collé deux colonnes. On reste volontairement étroit : une virgule ou
# un slash appartiennent souvent à la référence elle-même (« 123.456 »).
SEPARATEURS = re.compile(r"\s*\+\s*|\s{2,}")


def morceaux(reference) -> list[str]:
    """Les identifiants distincts que porte une cellule."""
    if reference is None or (isinstance(reference, float) and pd.isna(reference)):
        return []
    texte = str(reference).strip()
    if not texte:
        return []
    parts = [p.strip() for p in SEPARATEURS.split(texte) if p and p.strip()]
    return parts or [texte]


def variantes_ean(valeur: str) -> list[tuple[str, str]]:
    """Les EAN plausibles que cette valeur pourrait désigner, avec leur preuve.

    Ne renvoie QUE des codes dont la clé de contrôle est correcte. Un candidat
    dont la clé ne tombe pas juste n'est pas un EAN : le proposer reviendrait à
    fabriquer un prix faux et silencieux.
    """
    texte = str(valeur or "").strip()
    if not texte:
        return []

    # Un code-barres ne contient QUE des chiffres. « 1234-56-7.0 » est une
    # référence fournisseur : en retirer les tirets pour obtenir huit chiffres,
    # puis leur recalculer une clé, ne démontre rien — cela fabrique un EAN qui
    # n'a jamais existé. On refuse donc tout ce qui n'est pas déjà numérique.
    if not texte.isdigit():
        return []
    chiffres = texte

    trouves: list[tuple[str, str]] = []

    def ajouter(code: str, preuve: str) -> None:
        if ean_est_valide(code) and all(code != c for c, _ in trouves):
            trouves.append((code, preuve))

    # Tel quel
    ajouter(chiffres, "EAN lu tel quel")

    # GTIN-14 d'unite de vente : l'indicateur de tete vaut zero, les treize
    # chiffres de l'EAN sont deja tous la. On enleve un chiffre de bourrage,
    # on n'en invente aucun — c'est pourquoi ce cas survit a la regle.
    if len(chiffres) == 14 and chiffres[0] == "0":
        ajouter(chiffres[1:], "GTIN-14 d'unité ramené en EAN-13")

    # RETIRE LE 16/09, sur decision : ON NE RECONSTRUIT PAS D'IDENTIFIANT.
    #
    # Trois reparations vivaient ici, toutes demontrees par la cle de
    # controle, toutes supprimees :
    #   - EAN tronque a 12 chiffres, cle recalculee ;
    #   - zero de tete perdu par Excel, restitue par zfill ;
    #   - cle fausse d'un chiffre, corrigee.
    #
    # L'argument etait qu'une cle de controle ne tombe juste qu'une fois sur
    # dix par hasard. Il est vrai et il ne suffit pas : sur des milliers de
    # references, une fois sur dix arrive souvent, et le prix pose alors est
    # faux ET muet. Un identifiant abime se signale, il ne se repare pas.
    #
    # Ce que devient un tel code : il reste sans prix, et le rapprochement
    # suivant — libelle, fabricant — s'en charge en PROPOSANT.

    return trouves


def candidats_identifiant(reference) -> list[tuple[str, str]]:
    """Tous les EAN démontrables derrière une référence, cellule découpée."""
    resultats: list[tuple[str, str]] = []
    parts = morceaux(reference)
    for i, part in enumerate(parts):
        prefixe = "" if len(parts) == 1 else f"cellule à {len(parts)} identifiants — "
        for code, preuve in variantes_ean(part):
            if all(code != c for c, _ in resultats):
                resultats.append((code, prefixe + preuve))
    return resultats


# --------------------------------------------------------------------------
# Socle de référence — signalé, jamais appliqué


MOTIF_SOCLE = re.compile(r"(\d{6,})")


def socle_reference(reference) -> str | None:
    """Le radical numérique d'au moins six chiffres, comme chez le
    responsable des prix.

    « 1234567HR » et « 1234567 » partagent le socle « 1234567 ». C'est un
    indice, pas une preuve : chez le Fournisseur J, « UG » et « 6C »
    distinguent des sous-gammes qui ne sont pas le même produit.
    """
    if reference is None or (isinstance(reference, float) and pd.isna(reference)):
        return None
    trouve = MOTIF_SOCLE.search(str(reference))
    return trouve.group(1) if trouve else None


# --------------------------------------------------------------------------
# Application au rapprochement


def rapprocher_par_identifiant(fusion: pd.DataFrame,
                               tarif: pd.DataFrame,
                               champs: list[str]) -> pd.DataFrame:
    """Cinquième passe : les EAN LUS, sur ce qui reste sans prix.

    Appelée depuis pa_completer.py après les quatre clés existantes. Depuis
    le 16/09 elle ne rapproche que sur des codes entièrement présents dans
    la cellule — cellule à plusieurs identifiants, GTIN-14 dépadé. Plus
    aucune reconstruction, donc plus aucun prix posé sur un chiffre inventé.
    """
    if "prix_achat_unitaire_ht" not in fusion.columns:
        return fusion
    absents = fusion["prix_achat_unitaire_ht"].isna()
    if not absents.any():
        return fusion

    # Index des EAN du tarif, tous ceux qui sont valides.
    index: dict[str, int] = {}
    for colonne in ("ean", "ref_fournisseur"):
        if colonne not in tarif.columns:
            continue
        for position, valeur in tarif[colonne].items():
            chiffres = re.sub(r"\D", "", str(valeur or ""))
            if ean_est_valide(chiffres):
                index.setdefault(chiffres, position)
    if not index:
        return fusion

    sources = []
    for colonne in ("ref_fournisseur_wms", "ref_fabricant"):
        if colonne in fusion.columns:
            sources.append(colonne)
    if not sources:
        return fusion

    trouves = 0
    for i in fusion.index[absents]:
        for colonne in sources:
            gagne = False
            for code, preuve in candidats_identifiant(fusion.at[i, colonne]):
                if code in index:
                    ligne = index[code]
                    for champ in champs:
                        if champ in tarif.columns:
                            fusion.at[i, champ] = tarif.at[ligne, champ]
                    fusion.at[i, "Clé de rapprochement"] = "EAN lu"
                    fusion.at[i, "preuve_identifiant"] = f"{colonne} : {preuve} -> {code}"
                    trouves += 1
                    gagne = True
                    break
            if gagne:
                break
    if trouves:
        print(f"  {trouves} rapprochés par EAN lu")
    return fusion


# --------------------------------------------------------------------------
# Diagnostic


def diagnostic(motif: str | None = None) -> None:
    """Ce que la réparation d'identifiants trouverait, sans rien appliquer."""
    from extracteurs.base import chemin_lisible

    chemin = RACINE / "sortie" / "3-achats" / "pa_etat_par_article.xlsx"
    if not chemin.exists():
        print("pa_etat_par_article.xlsx absent : lancer run_pa.py d'abord")
        return
    d = pd.read_excel(chemin_lisible(chemin), sheet_name=0, header=3, dtype=str)
    if motif:
        d = d[d["Fournisseur"].fillna("").str.upper().str.contains(motif.upper())]

    cible = d[d["Motif"].fillna("").str.contains("pas dans son tarif", case=False)]
    print(f"{len(cible)} articles au tarif de leur fournisseur mais non rapprochés")
    print()

    compteurs: dict[str, int] = {}
    exemples: dict[str, list[str]] = {}
    multi = 0
    for _, r in cible.iterrows():
        for colonne in ("Réf. article fournisseur", "Référence fabricant"):
            if colonne not in cible.columns:
                continue
            brut = r.get(colonne)
            if len(morceaux(brut)) > 1:
                multi += 1
            for code, preuve in candidats_identifiant(brut):
                if preuve.endswith("EAN lu tel quel"):
                    continue
                cle = preuve.split(" — ")[-1]
                compteurs[cle] = compteurs.get(cle, 0) + 1
                exemples.setdefault(cle, []).append(f"{brut} -> {code}")

    print("=== reconstructions possibles ===")
    if compteurs:
        for cle, n in sorted(compteurs.items(), key=lambda x: -x[1]):
            print(f"  {n:>4}  {cle}")
            for e in exemples[cle][:4]:
                print(f"          {e}")
    else:
        print("  aucune")
    print(f"\n  cellules à identifiants multiples : {multi}")

    # Le socle : on compte, on n'applique pas.
    socles = cible["Réf. article fournisseur"].map(socle_reference).dropna()
    print(f"\n=== socles de référence exploitables : {len(socles)} ===")
    print("  (signalés, jamais appliqués — corroboration par le prix requise)")


def main() -> None:
    diagnostic(sys.argv[1] if len(sys.argv) > 1 else None)


if __name__ == "__main__":
    main()
