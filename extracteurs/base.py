# -*- coding: utf-8 -*-
"""
Socle commun a tous les extracteurs de catalogues fournisseurs.

Contient :
  - le schema de sortie normalise (COLONNES)
  - les helpers de nettoyage (texte, nombre, EAN)
  - la validation de la cle de controle EAN-8 / EAN-13
  - la detection automatique de la ligne d'en-tete dans un onglet Excel

Chaque extracteur (extracteurs/<fournisseur>.py) doit exposer :
    extract(path) -> pandas.DataFrame
respectant exactement COLONNES, dans cet ordre.
"""

from __future__ import annotations

import re
import shutil
import tempfile
import unicodedata
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Schema de sortie normalise, identique pour tous les fournisseurs
# ---------------------------------------------------------------------------

# Les noms sont explicites a dessein : un prix dit toujours s'il est
# unitaire ou par colis, un taux dit toujours qu'il est en decimal.
# L'ordre compte : les colonnes utiles au quotidien viennent en premier.
COLONNES = [
    "fournisseur",         # nom du fournisseur (ex: FOURNISSEUR A)
    "ref_fournisseur",     # reference catalogue fournisseur (TOUJOURS du texte)
    # Notre propre code article, quand le fichier fournisseur le porte.
    # C'est rare — seuls les fichiers issus d'un rapprochement anterieur
    # l'ont — mais c'est le lien le plus sur avec le WMS : il dispense de
    # passer par la reference fournisseur, qui est justement ce qui
    # diverge le plus souvent.
    "code_article_wms",
    "designation",         # libelle produit
    "ean",                 # code EAN de l'UNITE DE VENTE (toujours du texte)
    # Beaucoup de tarifs publient deux codes : celui de la piece et celui
    # du colis (Fournisseur AE "Unite"/"CDT", Fournisseur T "sachet"/"carton",
    # Fournisseur CV va jusqu'a trois : unite, boite, carton). `ean` porte toujours
    # l'unite de vente — c'est elle qu'on scanne au picking ; le code du
    # colis est garde a part, il sert a la reception.
    "ean_conditionnement",
    # --- prix d'achat -------------------------------------------------------
    "prix_achat_unitaire_ht",  # prix d'achat d'UNE unite, HT
    # Intitule EXACT de la colonne d'ou sort ce prix, tel qu'il est ecrit
    # dans le tarif. Un tarif publie souvent trois prix cote a cote —
    # public, remise, au colis — et se tromper de colonne ne provoque
    # aucune erreur : juste un PA faux. C'est la seule facon de verifier
    # apres coup qu'on a pris le bon.
    "colonne_prix",
    "conditionnement",         # nombre d'unites par colis
    "prix_colis_ht",           # prix_achat_unitaire_ht x conditionnement
    "conditionnement_suppose",  # True si le conditionnement etait absent -> 1
    # --- construction du prix, pour le controle de marge --------------------
    "tarif_public_ttc",    # tarif public conseille, TTC (jamais HT)
    "tva_taux",            # taux de TVA en DECIMAL (0.2 = 20 %)
    "remise_taux",         # taux de remise en DECIMAL, applique sur le HT
    "prix_recalcule",      # (tarif_public_ttc / (1 + tva)) x (1 - remise)
    "ecart_prix",          # |prix_recalcule - prix_achat_unitaire_ht|
    "prix_coherent",       # True si l'ecart est inferieur a TOLERANCE_PRIX
    # --- paliers degressifs : ce sont des PRIX UNITAIRES, pas des prix de lot
    "palier2_qte",         # a partir de cette quantite, le prix unitaire...
    "palier2_prix_ht",     # ... tombe a cette valeur
    "economie_palier2_pct",  # economie unitaire vs prix de base, en decimal
    "palier3_qte",
    "palier3_prix_ht",
    "economie_palier3_pct",
    "paliers",             # resume lisible : "1:30.40 | 6:28.70 | 75:17.60"
    # --- caracteristiques produit ------------------------------------------
    "eco_part_ht",         # eco-participation HT
    "code_lppr",           # code LPPR si present (texte)
    "montant_lppr",        # montant LPPR si present
    "dispositif_medical",  # O / N
    "origine",             # pays d'origine
    "fichier_source",      # nom du fichier d'ou vient la ligne
]

# Ecart tolere entre le prix net du fichier et le prix recalcule (en euros).
TOLERANCE_PRIX = 0.05

# Colonnes qui doivent rester du texte a l'ecriture Excel
COLONNES_TEXTE = [
    "fournisseur",
    "ref_fournisseur",
    "code_article_wms",
    "designation",
    "ean",
    "ean_conditionnement",
    "colonne_prix",
    "paliers",
    "code_lppr",
    "dispositif_medical",
    "origine",
    "fichier_source",
]

# Colonnes numeriques (float). Aucun arrondi n'est applique ici : les
# montants restent en pleine precision et ne sont arrondis qu'a l'affichage,
# par le format de cellule Excel.
COLONNES_NUM = [
    "prix_achat_unitaire_ht",
    "conditionnement",
    "prix_colis_ht",
    "tarif_public_ttc",
    "tva_taux",
    "remise_taux",
    "prix_recalcule",
    "ecart_prix",
    "palier2_qte",
    "palier2_prix_ht",
    "economie_palier2_pct",
    "palier3_qte",
    "palier3_prix_ht",
    "economie_palier3_pct",
    "eco_part_ht",
    "montant_lppr",
]

# Colonnes booleennes
COLONNES_BOOL = ["prix_coherent", "conditionnement_suppose"]


# ---------------------------------------------------------------------------
# Lecture des fichiers sources
# ---------------------------------------------------------------------------

def chemin_lisible(path) -> Path:
    """Chemin exploitable pour lire un classeur, meme ouvert dans Excel.

    Excel pose un verrou exclusif sur les fichiers ouverts : toute lecture
    echoue alors en PermissionError, ce qui interrompt tout le traitement
    pour une raison sans rapport avec les donnees. La copie, elle, reste
    autorisee : on travaille donc sur un double temporaire.

    Renvoie le chemin d'origine quand il est lisible directement.
    """
    path = Path(path)
    try:
        with open(path, "rb"):
            return path
    except PermissionError:
        pass

    copie = Path(tempfile.gettempdir()) / f"catalogue_verrouille_{path.name}"
    shutil.copy2(path, copie)
    print(f"    (fichier ouvert dans Excel, lecture d'une copie temporaire)")
    return copie


# ---------------------------------------------------------------------------
# Helpers de nettoyage
# ---------------------------------------------------------------------------

def nettoyer_texte(valeur) -> str | None:
    """Convertit n'importe quelle cellule en texte propre, ou None si vide.

    Point important : les references et les EAN arrivent parfois d'Excel sous
    forme de float (812176.0, 3.760123e+12). On repasse en entier avant de
    stringifier pour ne pas trainer un '.0' ou une notation scientifique.
    """
    if valeur is None:
        return None
    if isinstance(valeur, float):
        # NaN
        if valeur != valeur:
            return None
        # float entier -> on enleve le .0 parasite
        if valeur.is_integer():
            return str(int(valeur))
        return repr(valeur)
    if isinstance(valeur, int):
        return str(valeur)
    texte = str(valeur).strip()
    # Espaces insecables et doubles espaces issus des exports Excel
    texte = texte.replace(" ", " ")
    texte = re.sub(r"\s+", " ", texte)
    return texte or None


def nettoyer_nombre(valeur) -> float | None:
    """Convertit une cellule en float, ou None si non convertible / vide.

    Gere le separateur decimal virgule au cas ou (certains catalogues
    exportent les prix en texte francais), les symboles monetaires, les
    espaces de milliers et les pourcentages ecrits '20 %' -> 0.2.
    """
    if valeur is None:
        return None
    if isinstance(valeur, bool):
        return None
    if isinstance(valeur, (int, float)):
        if isinstance(valeur, float) and valeur != valeur:  # NaN
            return None
        return float(valeur)

    texte = str(valeur).strip()
    if not texte:
        return None

    pourcentage = "%" in texte
    # On ne garde que ce qui peut composer un nombre
    texte = texte.replace(" ", "").replace(" ", "")
    texte = texte.replace("€", "").replace("%", "")
    texte = texte.replace(",", ".")
    texte = re.sub(r"[^0-9.\-]", "", texte)
    if texte in ("", "-", ".", "-."):
        return None
    try:
        nombre = float(texte)
    except ValueError:
        return None
    return nombre / 100 if pourcentage else nombre


def nettoyer_ean(valeur) -> str | None:
    """Normalise un code EAN en TEXTE.

    - passe par nettoyer_texte pour tuer la notation scientifique
    - retire tout ce qui n'est pas un chiffre (tirets, espaces)
    - recomplete les zeros de tete perdus par Excel :
        11 ou 12 chiffres -> 13 (cas des UPC-A americains, dont le ou les
        zeros de tete sautent systematiquement a l'export)
        7 chiffres        -> 8
      Le zero-padding n'est applique que si la cle de controle devient
      correcte, sinon on garde la valeur telle quelle pour ne pas masquer
      une vraie erreur de saisie.
    - ramene a l'EAN-13 les GTIN-14 dont l'indicateur de colisage est 0 :
      un tel code ne designe pas un carton mais l'unite de vente, ecrite
      sur quatorze positions. La cle est la meme de part et d'autre, le
      zero de tete ne pesant rien dans la somme ponderee.
    - ecarte ce qui est trop court pour etre un code-barres (cellules a
      "0", restes de mise en page)
    Ne fait AUCUN filtrage de validite : elle est evaluee separement par
    ean_est_valide() pour pouvoir la signaler sans supprimer la ligne.
    """
    texte = nettoyer_texte(valeur)
    if texte is None:
        return None
    chiffres = re.sub(r"\D", "", texte)
    if not chiffres:
        return None

    if len(chiffres) in (11, 12):
        candidat = chiffres.zfill(13)
        if cle_controle_ean(candidat) == int(candidat[-1]):
            return candidat
    elif len(chiffres) == 7:
        candidat = chiffres.zfill(8)
        if cle_controle_ean(candidat) == int(candidat[-1]):
            return candidat
    elif len(chiffres) == 14 and chiffres[0] == "0":
        # Indicateur 0 : unite de vente, pas colisage
        candidat = chiffres[1:]
        if cle_controle_ean(candidat) == int(candidat[-1]):
            return candidat

    # En dessous de huit chiffres, aucun code-barres n'existe : c'est une
    # cellule vide ou un residu, pas une donnee a signaler.
    if len(chiffres) < 8:
        return None

    return chiffres


def cle_controle_ean(chiffres: str) -> int | None:
    """Calcule la cle de controle attendue pour un EAN-8 ou EAN-13.

    Algorithme standard GS1 : somme ponderee 3/1 en partant de la droite
    (hors cle), puis complement a la dizaine superieure.
    """
    if not chiffres or not chiffres.isdigit() or len(chiffres) not in (8, 12, 13, 14):
        return None
    corps = chiffres[:-1]
    somme = 0
    # Le poids alterne 3 puis 1 en remontant depuis le dernier chiffre du corps
    for position, caractere in enumerate(reversed(corps)):
        poids = 3 if position % 2 == 0 else 1
        somme += int(caractere) * poids
    return (10 - somme % 10) % 10


def ean_est_valide(ean) -> bool:
    """True si l'EAN a une longueur 8 ou 13 ET une cle de controle correcte.

    Un GTIN-14 (code de colisage) est volontairement considere comme NON
    valide ici : il est syntaxiquement correct mais ne designe pas l'unite
    de vente, ce n'est donc pas ce qu'on veut injecter dans le WMS.
    diagnostic_ean() permet de le distinguer d'une vraie erreur.
    """
    if not ean or not isinstance(ean, str):
        return False
    if not ean.isdigit() or len(ean) not in (8, 13):
        return False
    # « 0000000000000 » passe la cle de controle : une somme nulle est
    # divisible par dix. C'est pourtant une case vide remplie au clavier,
    # pas un code — le tarif du Fournisseur I en porte 48, dont 34 auraient
    # ECRASE un vrai code dans le WMS. La cle prouve la coherence d'un
    # nombre, jamais qu'il designe un produit.
    if set(ean) == {"0"}:
        return False
    return cle_controle_ean(ean) == int(ean[-1])


def diagnostic_ean(ean) -> str:
    """Explique pourquoi un EAN est refuse, pour rendre le rapport actionnable.

    Renvoie une chaine courte parmi :
      "OK"                      -> EAN-8 / EAN-13 avec cle correcte
      "absent"
      "caracteres non numeriques"
      "GTIN-14 (colisage), cle correcte" -> code carton, pas l'unite de vente
      "longueur N inattendue"
      "cle de controle incorrecte (attendue X)"
    """
    if ean is None or not isinstance(ean, str) or ean == "":
        return "absent"
    if not ean.isdigit():
        return "caracteres non numeriques"

    longueur = len(ean)
    cle_attendue = cle_controle_ean(ean)

    if longueur == 14:
        if cle_attendue != int(ean[-1]):
            return "longueur 14 et cle incorrecte"
        # L'indicateur de tete distingue le carton de l'unite de vente
        if ean[0] == "0":
            return "GTIN-14 d'unite (a ramener en EAN-13)"
        return "GTIN-14 (colisage), cle correcte"
    if longueur == 12:
        if cle_attendue == int(ean[-1]):
            return "UPC-A 12 (a completer en EAN-13)"
        return "longueur 12 et cle incorrecte"
    if longueur not in (8, 13):
        return f"longueur {longueur} inattendue"
    if cle_attendue != int(ean[-1]):
        return f"cle de controle incorrecte (attendue {cle_attendue})"
    return "OK"


# ---------------------------------------------------------------------------
# Logique tarifaire
# ---------------------------------------------------------------------------
#
# Chaine de calcul du tarif fournisseur, verifiee sur le tarif 2026 du
# Fournisseur A :
#
#     prix_ht_public = tarif_public_ttc / (1 + tva_taux)
#     prix_achat     = prix_ht_public * (1 - remise_taux)
#
# Exemples : 480,00 / 1,20 = 400,00 ; x (1 - 0,30000) = 280,00   # valeurs d'exemple
#            527,50 / 1,055 = 500,00 ; x (1 - 0,40000) = 300,00  # valeurs d'exemple
#
# Le tarif public de la colonne 3 est TTC, jamais HT. C'est la source
# d'erreur numero un sur ce fichier.

def normaliser_conditionnement(valeur) -> tuple[float, bool]:
    """Conditionnement exploitable, et indicateur de valeur supposee.

    Un conditionnement vide, nul, negatif ou non numerique est ramene a 1
    (on considere alors que le produit se commande a l'unite), et le second
    element du retour vaut True pour garder la trace de cette hypothese.
    """
    nombre = nettoyer_nombre(valeur)
    if nombre is None or nombre <= 0:
        return 1.0, True
    return float(nombre), False


def calculer_prix_colis(prix_unitaire_ht, conditionnement) -> float | None:
    """Prix d'un colis complet : prix unitaire x nombre d'unites par colis.

    Aucun arrondi : la valeur reste en pleine precision, l'arrondi est du
    ressort du format d'affichage Excel.
    """
    if prix_unitaire_ht is None or conditionnement is None:
        return None
    return prix_unitaire_ht * conditionnement


def controler_coherence_prix(
    tarif_public_ttc, tva_taux, remise_taux, prix_achat_unitaire_ht
) -> tuple[float | None, float | None, bool | None]:
    """Recalcule le prix d'achat theorique et le compare au prix du fichier.

    Renvoie (prix_recalcule, ecart absolu, coherent). Un ecart superieur a
    TOLERANCE_PRIX signale soit une erreur du fournisseur, soit un cas
    particulier (prix negocie hors grille) : dans les deux cas la ligne doit
    etre verifiee avant de servir de base a une commande.

    Renvoie (None, None, None) quand une des valeurs necessaires manque :
    on ne peut alors ni confirmer ni infirmer la coherence.
    """
    if (
        tarif_public_ttc is None
        or tva_taux is None
        or remise_taux is None
        or prix_achat_unitaire_ht is None
        or tva_taux <= -1
    ):
        return None, None, None

    prix_recalcule = (tarif_public_ttc / (1 + tva_taux)) * (1 - remise_taux)
    ecart = abs(prix_recalcule - prix_achat_unitaire_ht)
    return prix_recalcule, ecart, ecart < TOLERANCE_PRIX


def valider_paliers(prix_unitaire, paliers) -> list:
    """Ne retient que les paliers qui en sont vraiment.

    Un palier degressif suppose deux choses : une quantite superieure a 1,
    et un prix inferieur au prix unitaire. Tout ce qui ne remplit pas ces
    conditions est une AUTRE colonne de prix — le plus souvent le prix
    public conseille, que le fournisseur publie a cote de son prix net.

    Confondre les deux ferait apparaitre des "remises" negatives et
    fausserait toute l'analyse d'achat par quantite.

    `paliers` est une liste de couples (quantite, prix). Le retour est la
    meme liste, filtree et triee par quantite croissante.
    """
    if prix_unitaire is None or prix_unitaire <= 0:
        return []

    retenus = [
        (quantite, prix)
        for quantite, prix in paliers
        if quantite is not None and prix is not None
        and quantite > 1 and 0 < prix < prix_unitaire
    ]
    return sorted(retenus, key=lambda couple: couple[0])


def economie_palier(prix_base, prix_palier) -> float | None:
    """Economie unitaire d'un palier par rapport au prix de base, en decimal.

    Une valeur negative signale un palier plus cher que le prix de base,
    c'est-a-dire une anomalie du tarif fournisseur.
    """
    if prix_base is None or prix_palier is None or prix_base <= 0:
        return None
    return 1 - prix_palier / prix_base


def resumer_paliers(prix_unitaire, paliers, quantite_base=1) -> str | None:
    """Resume lisible de la grille degressive : "1:30.40 | 6:28.70 | 75:17.60".

    `paliers` est une liste de couples (quantite seuil, prix unitaire a ce
    palier). Format destine a la lecture humaine uniquement : les colonnes
    palier2_qte / palier2_prix_ht restent la source pour tout calcul.

    `quantite_base` est la quantite a partir de laquelle le prix unitaire
    s'applique. Elle vaut 1 dans la plupart des tarifs, mais certains ne
    cotent qu'a partir d'un lot (le Fournisseur L demarre a 20 sur des
    cannes) :
    afficher "1:" serait alors trompeur.
    """
    if prix_unitaire is None:
        return None

    valides = [
        (quantite, prix)
        for quantite, prix in paliers
        if quantite is not None and prix is not None
    ]
    if not valides:
        return None

    valides.sort(key=lambda couple: couple[0])
    morceaux = [f"{quantite_base:g}:{prix_unitaire:.2f}"]
    morceaux += [f"{quantite:g}:{prix:.2f}" for quantite, prix in valides]
    return " | ".join(morceaux)


def anomalies_paliers(prix_base, palier2, palier3) -> list[str]:
    """Verifie la coherence de la grille degressive.

    `palier2` et `palier3` sont des couples (quantite, prix unitaire).
    Une grille saine est strictement decroissante en prix quand la quantite
    augmente. Renvoie la liste des anomalies constatees, vide si tout va bien.
    """
    anomalies = []
    q2, p2 = palier2
    q3, p3 = palier3

    if prix_base is not None:
        for nom, prix in (("palier 2", p2), ("palier 3", p3)):
            if prix is not None and prix > prix_base:
                anomalies.append(f"{nom} plus cher que le prix de base")

    if q2 is not None and q3 is not None:
        if q3 <= q2:
            anomalies.append("quantite du palier 3 inferieure ou egale au palier 2")
            if p2 is not None and p3 is not None and p3 < p2:
                anomalies.append("palier 3 moins cher a quantite inferieure")
        elif p2 is not None and p3 is not None and p3 > p2:
            anomalies.append("palier 3 plus cher que le palier 2")

    return anomalies


def prix_pour_quantite(ligne, quantite: int) -> dict:
    """Meilleur tarif applicable pour une quantite donnee.

    `ligne` est une ligne du catalogue (dict ou Series pandas) au schema
    commun. La quantite demandee est arrondie au multiple superieur du
    conditionnement : on ne peut pas commander 3 unites d'un produit vendu
    par 6. Le palier retenu est ensuite choisi d'apres la quantite
    reellement commandee, pas d'apres celle demandee.

    Renvoie :
        prix_unitaire_ht    prix unitaire du palier applique
        total_ht            prix_unitaire_ht x quantite_commandee
        palier_applique     "base", "palier 2" ou "palier 3"
        quantite_demandee   la quantite passee en argument
        quantite_commandee  apres arrondi au conditionnement
        arrondi_conditionnement  True si la quantite a du etre relevee
        conditionnement     conditionnement retenu
    """
    def valeur(champ):
        """Lecture tolerante : dict, Series pandas, valeurs NaN."""
        brut = ligne[champ] if champ in ligne else None
        if brut is None or (isinstance(brut, float) and brut != brut):
            return None
        return brut

    if quantite is None or quantite <= 0:
        raise ValueError("La quantite doit etre un entier strictement positif")

    prix_base = valeur("prix_achat_unitaire_ht")
    if prix_base is None:
        raise ValueError(
            f"Prix unitaire absent pour la reference {valeur('ref_fournisseur')}"
        )

    conditionnement, _ = normaliser_conditionnement(valeur("conditionnement"))

    # On ne commande que des colis entiers
    nb_colis = -(-quantite // conditionnement)  # division entiere par exces
    quantite_commandee = nb_colis * conditionnement
    arrondi = quantite_commandee != quantite

    # Le meilleur palier atteint par la quantite reellement commandee.
    # On parcourt dans l'ordre croissant des seuils et on garde le dernier
    # atteint, ce qui reste correct meme si la grille est mal ordonnee.
    prix_unitaire = prix_base
    palier_applique = "base"
    candidats = [
        ("palier 2", valeur("palier2_qte"), valeur("palier2_prix_ht")),
        ("palier 3", valeur("palier3_qte"), valeur("palier3_prix_ht")),
    ]
    for nom, seuil, prix in sorted(
        (c for c in candidats if c[1] is not None and c[2] is not None),
        key=lambda c: c[1],
    ):
        if quantite_commandee >= seuil and prix < prix_unitaire:
            prix_unitaire = prix
            palier_applique = nom

    return {
        "prix_unitaire_ht": prix_unitaire,
        "total_ht": prix_unitaire * quantite_commandee,
        "palier_applique": palier_applique,
        "quantite_demandee": quantite,
        "quantite_commandee": quantite_commandee,
        "arrondi_conditionnement": arrondi,
        "conditionnement": conditionnement,
    }


# ---------------------------------------------------------------------------
# Normalisation des references, pour le rapprochement avec la base du WMS
# ---------------------------------------------------------------------------

def cle_ref_stricte(valeur) -> str | None:
    """Cle de rapprochement stricte : majuscules, sans espaces.

    Le '.0' final est retire : il apparait des qu'une reference purement
    numerique a transite par un type flottant.
    """
    texte = nettoyer_texte(valeur)
    if texte is None:
        return None
    texte = texte.upper().replace(" ", "")
    if texte.endswith(".0"):
        texte = texte[:-2]
    return texte or None


def cle_ref_souple(valeur) -> str | None:
    """Cle de rapprochement souple : en plus, sans separateurs . - _ /

    Indispensable ici : le WMS stocke les declinaisons avec un point
    (123456.M, 789012.B) la ou le catalogue fournisseur les colle
    (123456M, 789012N). Cette seule tolerance fait gagner ~250
    rapprochements sur le Fournisseur A.
    """
    texte = cle_ref_stricte(valeur)
    if texte is None:
        return None
    texte = re.sub(r"[.\-_/]", "", texte)
    return texte or None


# ---------------------------------------------------------------------------
# Detection de la ligne d'en-tete
# ---------------------------------------------------------------------------

def _normaliser(texte: str) -> str:
    """Minuscule sans accent ni ponctuation, pour comparer des en-tetes."""
    texte = unicodedata.normalize("NFKD", texte)
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", texte.lower())


def trouver_entete_df(brut, mots_cles, max_lignes: int = 60) -> int | None:
    """Meme recherche que trouver_ligne_entete, mais sur un DataFrame brut.

    Renvoie l'index (0-based, pandas) de la premiere ligne contenant tous
    les mots-cles, ou None. Utile quand le fichier a ete lu avec
    `header=None` plutot qu'ouvert via openpyxl.
    """
    cles = [_normaliser(m) for m in mots_cles]
    for position in range(min(max_lignes, len(brut))):
        cellules = [
            _normaliser(str(valeur))
            for valeur in brut.iloc[position].tolist()
            if str(valeur) != "nan"
        ]
        if all(any(cle in cellule for cellule in cellules) for cle in cles):
            return position
    return None


def colonne_par_intitule(entetes, *fragments: str) -> int | None:
    """Position de la premiere colonne dont l'intitule contient un fragment.

    Les intitules des tarifs sont souvent sur plusieurs lignes et truffes
    de fautes ("Descirption"), d'ou la comparaison sur fragment normalise
    plutot que sur egalite.
    """
    cibles = [_normaliser(fragment) for fragment in fragments]
    for position, intitule in enumerate(entetes):
        texte = _normaliser(str(intitule))
        if any(cible in texte for cible in cibles):
            return position
    return None


def trouver_ligne_entete(feuille, mots_cles, max_lignes: int = 60) -> int:
    """Renvoie le numero (1-based, openpyxl) de la ligne d'en-tete.

    On cherche la PREMIERE ligne qui contient tous les mots-cles donnes
    (comparaison insensible a la casse et aux accents). Cela evite de coder
    en dur un offset qui differera d'un catalogue a l'autre.

    Leve ValueError si rien n'est trouve dans les `max_lignes` premieres.
    """
    cles = [_normaliser(m) for m in mots_cles]
    for ligne in feuille.iter_rows(min_row=1, max_row=max_lignes, values_only=False):
        cellules = [_normaliser(str(c.value)) for c in ligne if c.value is not None]
        if all(any(cle in cellule for cellule in cellules) for cle in cles):
            return ligne[0].row
    raise ValueError(
        f"Ligne d'en-tete introuvable (mots-cles cherches : {mots_cles})"
    )


# ---------------------------------------------------------------------------
# Finalisation du DataFrame
# ---------------------------------------------------------------------------

def finaliser(lignes: list[dict]) -> pd.DataFrame:
    """Construit le DataFrame final : colonnes dans l'ordre + bons types.

    Les colonnes texte sont forcees en 'object' avec des vraies chaines,
    jamais des nombres, pour que l'ecriture Excel conserve les zeros de tete
    et les references alphanumeriques (ex: 123456N).
    """
    df = pd.DataFrame(lignes, columns=COLONNES)

    for colonne in COLONNES_TEXTE:
        df[colonne] = df[colonne].astype("object").where(df[colonne].notna(), None)

    for colonne in COLONNES_NUM:
        df[colonne] = pd.to_numeric(df[colonne], errors="coerce")

    # Les booleens restent en 'object' : prix_coherent vaut None quand la
    # coherence n'a pas pu etre evaluee, ce qu'un dtype bool ecraserait.
    for colonne in COLONNES_BOOL:
        df[colonne] = df[colonne].astype("object")

    return df
