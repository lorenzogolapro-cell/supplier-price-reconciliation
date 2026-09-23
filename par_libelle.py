# -*- coding: utf-8 -*-
"""
Retrouve un produit par son libelle quand sa reference ne suffit pas.

La reference fournisseur fait foi, mais elle n'est pas toujours juste
dans le WMS : chez le Fournisseur S, les fauteuils de la GAMME EXEMPLE 1
portent une reference qui ne
suit pas celle du fabricant, alors que leur libelle, lui, est fiable. Le
libelle sert donc a deux choses :

  - RATTRAPER  quand aucune reference ne correspond, chercher le produit
               par son nom, a fournisseur egal ;
  - TRANCHER   quand une reference correspond mais que les deux libelles
               n'ont rien de commun, s'en mefier.

La methode est celle qui a fait ses preuves sur VIDAL : on compte la part
de NOS mots retrouvee chez le fournisseur — et non la ressemblance
mutuelle, car les catalogues decrivent plus longuement que nos libelles —
et on exige que les nombres concordent. Ce dernier point est le garde-fou
essentiel : sans lui, une taille L recevrait le code d'une taille S.

Un code obtenu ainsi est marque "libellé" : c'est une piste solide, pas
une certitude, et il ne doit pas partir dans le WMS sans relecture.
"""

from __future__ import annotations

import re
import unicodedata

import pandas as pd

# Vocabulaire trop courant pour distinguer deux produits
MOTS_VIDES = {
    "DE", "DU", "LA", "LE", "LES", "ET", "AVEC", "SANS", "POUR", "PAR",
    "EN", "UN", "UNE", "DES", "SUR", "AU", "AUX",
    "CM", "MM", "ML", "KG", "CL", "GR",
    "BTE", "BOITE", "SACHET", "CARTON", "PCS", "PIECE", "PIECES", "UNITE",
    "LOT", "PAIRE", "TAILLE", "POINTURE", "COLORIS", "REF", "MODELE",
}

# Longueur de racine : absorbe les pluriels et nos libelles tronques
RACINE = 5

# Part de nos mots qui doit se retrouver chez le fournisseur
SEUIL = 0.6

# Tous les nombres, meme colles a des lettres : "T39", "CH12", "0,8cm"
NOMBRE = re.compile(r"\d+(?:[.,]\d+)?")

# Mots qui distinguent deux produits d'une meme gamme. Presents d'un seul
# cote, ils disqualifient le rapprochement — c'est exactement ce qui
# manquait quand "MODELE EXEMPLE 1 LIGHT" a recu le code de
# "MODELE EXEMPLE 1 EXTRA Light", et que vingt fauteuils de la GAMME
# EXEMPLE 1 ont herite du code d'un modele bariatrique.
DISCRIMINANTS = {
    # intensite / gamme
    "EXTRA", "SUPER", "ULTRA", "MAXI", "MINI", "PLUS", "PREMIUM", "LIGHT",
    "NORMAL", "STANDARD", "CONFORT", "BASIC", "PRO", "EVO", "BARIATRIQUE",
    # morphologie
    "XXL", "XXS", "JUNIOR", "ENFANT", "ADULTE", "PEDIATRIQUE", "BEBE",
    "COURT", "LONG", "LARGE", "ETROIT", "HAUT", "BAS",
    # tailles d'une ou deux lettres : trop courtes pour etre retenues
    # comme mots, mais ce sont elles qui separent le S du L
    "S", "M", "L", "XL", "XS", "TS", "TM", "TL", "TXL", "TU",
    # laterite et forme
    "GAUCHE", "DROITE", "DROIT", "PLIANT", "FIXE", "ELECTRIQUE", "MANUEL",
    # couleurs : deux coloris sont deux articles
    "NOIR", "NOIRE", "BLANC", "BLANCHE", "BLEU", "BLEUE", "ROUGE", "VERT",
    "VERTE", "GRIS", "GRISE", "BEIGE", "TAUPE", "MARRON", "ROSE", "JAUNE",
    "ORANGE", "VIOLET", "ANTHRACITE", "CHOCO", "GREGE", "CHINE", "SAPHIR",
}

# Quantite par conditionnement : "SACHET DE 24", "BTE/30", "LOT DE 5",
# "x 12". Deux conditionnements differents sont deux articles — c'est le
# cas du sachet de 24 recu par le sachet de 14.
QUANTITE = re.compile(
    r"(?:SACHET|BOITE|BTE|LOT|PAQUET|PACK|CARTON|BLISTER|SET|POCHE)"
    r"[\s/]*(?:DE\s*)?(\d{1,4})\b|"
    r"\b(?:X|PAR)\s*(\d{1,4})\b|"
    r"\b(\d{1,4})\s*(?:PCS|PIECES?|UNITES?|U\.)\b",
    re.IGNORECASE,
)


def _normaliser(texte) -> str:
    if texte is None or (isinstance(texte, float) and texte != texte):
        return ""
    sans_accent = unicodedata.normalize("NFKD", str(texte).upper())
    sans_accent = "".join(c for c in sans_accent
                          if not unicodedata.combining(c))
    # L'apostrophe d'une contraction ne separe pas deux mots : « MODEL'S »
    # vaut UN mot, pas « MODEL » + « S ». La retirer AVANT de remplacer le
    # reste par des espaces evite de fabriquer un « S » parasite — qui est
    # aussi l'abreviation de la taille Small dans DISCRIMINANTS. Sans ce
    # retrait, deux produits « MODEL'S ... » PARTAGENT TOUJOURS ce « S », et
    # le controle de discriminants (couleur, cote) ne voit plus leurs vrais
    # desaccords : « MODEL'S GO NOIR » passait comme compatible avec
    # « Model's Go Beige », decouvert le 17/09 sur le lot du Fournisseur AN.
    sans_apostrophe = re.sub(r"['’]", "", sans_accent)
    return re.sub(r"[^A-Z0-9]+", " ", sans_apostrophe)


def mots(texte, marque: str | None = None) -> set:
    """Mots significatifs d'un libelle, marque et vocabulaire courant otes.

    Nos libelles commencent par la marque ("FOURNISSEUR S GAMME EXEMPLE 1...") que le
    catalogue du fabricant ne repete pas : la compter fausserait le score
    dans les deux sens.
    """
    normalise = _normaliser(texte)
    prefixe = _normaliser(marque)
    if prefixe and normalise.startswith(prefixe):
        normalise = normalise[len(prefixe):]

    retenus = set()
    for mot in normalise.split():
        if len(mot) < 3 or mot in MOTS_VIDES or mot.isdigit():
            continue
        retenus.add(mot[:RACINE])
    return retenus


def nombres(texte) -> set:
    """Nombres du libelle : tailles, calibres, dimensions.

    Deux produits dont les nombres different ne sont pas le meme, quelle
    que soit la ressemblance des mots.
    """
    valeurs = set()
    for brut in NOMBRE.findall(str(texte or "")):
        normalise = brut.replace(",", ".").rstrip("0").rstrip(".")
        valeurs.add(normalise.lstrip("0") or "0")
    return valeurs


def discriminants(texte, marque: str | None = None) -> set:
    """Mots qui distinguent deux produits d'une meme gamme."""
    normalise = _normaliser(texte)
    prefixe = _normaliser(marque)
    if prefixe and normalise.startswith(prefixe):
        normalise = normalise[len(prefixe):]
    return {mot for mot in normalise.split() if mot in DISCRIMINANTS}


def quantite(texte) -> set:
    """Quantites par conditionnement citees dans le libelle."""
    valeurs = set()
    for groupes in QUANTITE.findall(str(texte or "")):
        for valeur in groupes:
            if valeur:
                valeurs.add(str(int(valeur)))
    return valeurs


def indexer(catalogue: pd.DataFrame, colonne_fournisseur: str = "fournisseur",
            colonne_libelle: str = "designation") -> dict:
    """Prepare le catalogue pour la recherche : un index par fournisseur."""
    index: dict[str, list] = {}
    for ligne in catalogue.itertuples():
        fournisseur = getattr(ligne, colonne_fournisseur, None)
        libelle = getattr(ligne, colonne_libelle, None)
        ean = getattr(ligne, "ean", None)
        if not fournisseur or not libelle or not ean:
            continue
        index.setdefault(str(fournisseur).upper(), []).append({
            "libelle": libelle,
            "ean": ean,
            "reference": getattr(ligne, "ref_fournisseur", None),
            "mots": mots(libelle),
            "nombres": nombres(libelle),
            "discriminants": discriminants(libelle),
            "quantite": quantite(libelle),
        })
    return index


def chercher(libelle: str, candidats: list[dict], marque: str | None = None,
             seuil: float = SEUIL) -> dict | None:
    """Meilleur produit du catalogue pour un de nos libelles."""
    nos_mots = mots(libelle, marque)
    nos_nombres = nombres(libelle)
    nos_discriminants = discriminants(libelle, marque)
    notre_quantite = quantite(libelle)
    if len(nos_mots) < 2:
        return None

    meilleur, score_max = None, 0.0
    for candidat in candidats:
        if not candidat["mots"]:
            continue
        # Les nombres priment : ils distinguent les declinaisons
        if nos_nombres and candidat["nombres"] and not (
            nos_nombres & candidat["nombres"]
        ):
            continue

        # Un mot discriminant present d'un seul cote separe deux produits
        # d'une meme gamme : "Extra Light" n'est pas "Light", un modele
        # bariatrique n'est pas le modele courant.
        if nos_discriminants ^ candidat["discriminants"]:
            continue

        # Deux conditionnements differents sont deux articles : le sachet
        # de 24 ne porte pas le code du sachet de 14.
        if notre_quantite and candidat["quantite"] and not (
            notre_quantite & candidat["quantite"]
        ):
            continue

        communs = nos_mots & candidat["mots"]
        if not communs:
            continue
        score = len(communs) / len(nos_mots)
        if score > score_max or (
            score == score_max and meilleur is not None
            and len(candidat["libelle"]) < len(meilleur["libelle"])
        ):
            meilleur, score_max = candidat, score

    if meilleur is None or score_max < seuil:
        return None
    return {**meilleur, "score": round(score_max, 2)}


def enrichir(df: pd.DataFrame, catalogue: pd.DataFrame,
             colonne_fournisseur: str = "nom_fournisseur_wms",
             colonne_libelle: str = "designation_wms",
             correspondance=None) -> pd.DataFrame:
    """Complete les EAN manquants par rapprochement de libelle.

    `correspondance` ramene le nom du fournisseur tel que le WMS le porte a
    celui sous lequel le catalogue publie ses lignes (le WMS dit
    "EXEMPLE' S.A.", le tarif dit "EXEMPLE SA").
    """
    df = df.copy()
    if "source_ean" not in df.columns:
        df["source_ean"] = None

    index = indexer(catalogue)
    a_chercher = df["ean"].isna()
    if not a_chercher.any():
        return df

    trouves = 0
    for position in df.index[a_chercher]:
        fournisseur = df.at[position, colonne_fournisseur]
        if not fournisseur or pd.isna(fournisseur):
            continue

        cible = correspondance(fournisseur) if correspondance else fournisseur
        candidats = index.get(str(cible).upper())
        if not candidats:
            continue

        resultat = chercher(df.at[position, colonne_libelle], candidats,
                            marque=fournisseur)
        if not resultat:
            continue
        df.at[position, "ean"] = resultat["ean"]
        df.at[position, "source_ean"] = "libellé"
        df.at[position, "libelle_fournisseur"] = resultat["libelle"]
        # Le score dit quelle part de notre libelle s'est retrouvee chez
        # le fournisseur : au-dessus de 0,8 le produit est presque
        # toujours le bon, en dessous il faut regarder.
        df.at[position, "concordance_libelle"] = resultat["score"]
        trouves += 1

    print(f"  EAN retrouvés par le libellé : {trouves} (à vérifier)")
    return purger_doublons(df)


def purger_doublons(df: pd.DataFrame, sources: tuple = ("libellé",)
                    ) -> pd.DataFrame:
    """Retire les codes deduits attribues a PLUSIEURS de nos articles.

    Un code EAN designe un produit et un seul. Quand un rapprochement
    par libelle donne le meme code a vingt fauteuils qui ne different que
    par la couleur et la largeur d'assise, dix-neuf sont faux — et rien
    ne dit lequel est le bon. On les retire tous : un manque se comble,
    un code faux fait scanner un produit pour un autre.

    Les codes issus d'un tarif ou d'une reference ne sont pas concernes :
    un fournisseur peut legitimement publier le meme code sur deux lignes
    de son catalogue.
    """
    deduits = df["source_ean"].isin(sources) & df["ean"].notna()
    if not deduits.any():
        return df

    comptes = df.loc[deduits, "ean"].value_counts()
    partages = set(comptes[comptes > 1].index)
    if not partages:
        return df

    a_purger = deduits & df["ean"].isin(partages)
    retires = int(a_purger.sum())
    df.loc[a_purger, ["ean", "source_ean", "libelle_fournisseur",
                      "concordance_libelle"]] = None
    print(f"  {retires} retirés : {len(partages)} codes attribués à "
          f"plusieurs articles (rapprochement trop large)")
    return df
