# -*- coding: utf-8 -*-
"""
Decode un code-barres GS1-128 / DataMatrix GS1 lu a la douchette.

Les etiquettes de dispositifs medicaux ne portent pas un simple EAN-13
mais une chaine structuree, ou chaque donnee est precedee d'un
identifiant d'application (AI) entre parentheses :

    (01)01234567890128(17)270329(10)L4521

    (01) GTIN      -> le code produit, 14 chiffres
    (17) 270329    -> peremption, 29 mars 2027
    (10) L4521     -> numero de lot

Ce qui nous interesse est le GTIN. Quand son premier chiffre — l'indicateur
de conditionnement — vaut 0, il s'agit de l'unite de vente et le GTIN est
un EAN-13 precede d'un zero de remplissage : c'est ce code a treize
chiffres qu'il faut retrouver dans le WMS. Un indicateur de 1 a 8 designe
un carton ou une palette, qui n'est pas l'article vendu.

    python decoder_gs1.py "(01)01234567890128(17)270329"

Utile aussi pour verifier un scan qui n'a rien donne en picking : si la
douchette transmet la chaine entiere au lieu du seul GTIN, le WMS ne
trouve rien alors que le code est bon.
"""

from __future__ import annotations

import re
import sys

from extracteurs.base import cle_controle_ean, ean_est_valide

# Longueur fixe des AI les plus courants. Ceux qui n'y figurent pas sont
# de longueur variable et se terminent au separateur ou a la parenthese
# suivante.
LONGUEURS_FIXES = {
    "00": 18,   # SSCC, unite logistique
    "01": 14,   # GTIN
    "02": 14,   # GTIN des articles contenus
    "11": 6,    # date de fabrication
    "13": 6,    # date de conditionnement
    "15": 6,    # date de durabilite minimale
    "17": 6,    # date de peremption
    "20": 2,    # variante produit
}

LIBELLES = {
    "00": "unité logistique (SSCC)",
    "01": "GTIN (code produit)",
    "02": "GTIN des articles contenus",
    "10": "numéro de lot",
    "11": "date de fabrication",
    "15": "à consommer de préférence avant",
    "17": "date de péremption",
    "21": "numéro de série",
    "30": "quantité",
    "37": "nombre d'articles contenus",
    "240": "référence fabricant",
    "241": "référence client",
}

# Les dates GS1 s'ecrivent AAMMJJ
DATE = re.compile(r"^(\d{2})(\d{2})(\d{2})$")


def _lisible(ai: str, valeur: str) -> str:
    """Valeur mise en forme selon la nature de l'AI."""
    if ai in ("11", "13", "15", "17"):
        trouve = DATE.match(valeur)
        if trouve:
            annee, mois, jour = trouve.groups()
            # GS1 : 00-49 -> 2000-2049, 50-99 -> 1950-1999
            siecle = "20" if int(annee) <= 49 else "19"
            jour = jour if jour != "00" else "dernier jour du mois"
            return f"{jour}/{mois}/{siecle}{annee}"
    return valeur


def decoder(chaine: str) -> dict:
    """Decompose une chaine GS1 en {AI: valeur}.

    Accepte les deux ecritures : avec parentheses, telle que l'affichent
    les etiquettes, ou brute avec separateurs FNC1, telle que l'emettent
    certaines douchettes.
    """
    if not chaine:
        return {}
    texte = chaine.strip()

    # Les lecteurs prefixent la chaine d'un identifiant de symbologie :
    # ]C1 pour un GS1-128, ]d2 pour un DataMatrix GS1, ]e0 pour un RSS.
    # Il annonce le type de code-barres et ne fait pas partie des donnees.
    texte = re.sub(r"^\][A-Za-z]\d", "", texte)

    # Forme parenthesee : le plus simple, les AI sont delimites
    if "(" in texte:
        elements = re.findall(r"\((\d{2,4})\)([^(]*)", texte)
        return {ai: valeur.strip() for ai, valeur in elements}

    # Forme brute : on avance en s'appuyant sur les longueurs fixes, et
    # on s'arrete au separateur pour les AI de longueur variable.
    texte = texte.replace("\x1d", "|")
    resultat: dict[str, str] = {}
    position = 0
    while position < len(texte) - 1:
        ai = texte[position:position + 2]
        position += 2
        longueur = LONGUEURS_FIXES.get(ai)
        if longueur:
            resultat[ai] = texte[position:position + longueur]
            position += longueur
        else:
            fin = texte.find("|", position)
            fin = len(texte) if fin == -1 else fin
            resultat[ai] = texte[position:fin]
            position = fin + 1
    return resultat


def ean_depuis_gs1(chaine: str) -> str | None:
    """EAN-13 de l'unite de vente contenue dans un code GS1, s'il y en a.

    Renvoie None quand le GTIN designe un carton ou une palette : ce
    n'est alors pas le code de l'article que nous vendons.
    """
    gtin = decoder(chaine).get("01")
    if not gtin:
        return None
    gtin = re.sub(r"\D", "", gtin)
    if len(gtin) == 13 and ean_est_valide(gtin):
        return gtin
    if len(gtin) != 14:
        return None
    if gtin[0] != "0":
        return None  # colisage, pas l'unite de vente
    candidat = gtin[1:]
    return candidat if ean_est_valide(candidat) else None


def expliquer(chaine: str) -> str:
    """Description lisible d'un code scanne."""
    elements = decoder(chaine)
    if not elements:
        return "chaîne vide ou illisible"

    lignes = []
    for ai, valeur in elements.items():
        libelle = LIBELLES.get(ai, f"AI {ai}")
        lignes.append(f"  ({ai}) {libelle:<32} {_lisible(ai, valeur)}")

    gtin = elements.get("01", "")
    if gtin:
        ean = ean_depuis_gs1(chaine)
        if ean:
            lignes.append(f"\n  -> EAN-13 à chercher dans le WMS : {ean}")
        elif len(gtin) == 14 and gtin[0] != "0":
            lignes.append(
                f"\n  -> indicateur {gtin[0]} : code de colisage "
                f"(carton ou palette), et non l'unité de vente"
            )
        else:
            attendue = cle_controle_ean(gtin)
            lignes.append(
                f"\n  -> GTIN invalide : clé lue {gtin[-1]}, "
                f"attendue {attendue}"
            )
    return "\n".join(lignes)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)
    for argument in sys.argv[1:]:
        print(f"\n{argument}")
        print(expliquer(argument))
