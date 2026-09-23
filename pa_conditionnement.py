# -*- coding: utf-8 -*-
"""
Chantier PA — ramener un prix au conditionnement vers un prix unitaire.

    python pa_conditionnement.py            diagnostic sur tout le périmètre
    python pa_conditionnement.py FOURNISSEUR_G  diagnostic sur un fournisseur

LE PROBLÈME
    Certains fournisseurs cotent au lot — par 50, par 100, la boîte, le carton,
    la paire — alors que le WMS raisonne à l'unité. Le prix lu est juste,
    mais il
    ne répond pas à la même question, et rien dans le fichier ne le dit.

    Un x2 chez le Fournisseur G n'est pas une hausse : c'est une chaussure
    vendue par paire quand le WMS compte les souliers un par un.

LE DIVISEUR, PAR ORDRE DE FIABILITÉ
    1. COLONNE   le tarif porte une colonne explicite — « Nbre unités par
                 boîte » chez le Fournisseur O, « PCB » chez le Fournisseur
                 AL, « CDT » chez le Fournisseur BH, « UNITE DE VENTE » chez
                 le Fournisseur BM. L'extracteur la range
                 dans `conditionnement` : elle fait foi.
    2. DÉSIGNATION  le libellé du tarif la porte — « 20X100ML », « BTE/50 »,
                 « LOT DE 25 ». Lecture de texte, donc motifs stricts : une
                 taille « T38 » ou un modèle « 1234567 » ne doivent jamais
                 passer pour un conditionnement.
    3. WMS       le conditionnement de la fiche article.
    4. DÉDUIT    le rapport au DPA tombe sur un entier rond.

    **Le rang 4 ne s'applique JAMAIS seul.** Une déduction par le rapport doit
    être corroborée, soit par une autre source, soit par sa systématicité : le
    même entier sur au moins cinq articles du même fournisseur. Un diviseur
    déduit sur un article isolé se SIGNALE, il ne s'applique pas.

    C'est le mur du Fournisseur BP : leur tarif mêle des articles cotés à
    l'unité et
    d'autres au conditionnement, sans indicateur. Le facteur n'y est pas
    déductible. Ne pas essayer de le forcer.

TROIS CONFIGURATIONS, À NE PAS CONFONDRE
    A. Le tarif donne plusieurs prix par quantité, dont un unitaire.
       On prend la colonne unitaire et **on ne divise rien**. Les autres
       colonnes sont de vrais paliers, déjà exprimés en unités.

    B. Le seul prix du tarif est un prix au conditionnement.
       On divise par N. Le « par 100 » n'est pas un palier, c'est l'unité de
       vente : il n'y a qu'un prix, donc aucun seuil à convertir. Le
       conditionnement devient un **minimum de commande**, rangé à part — il
       sert au réapprovisionnement, où le MOQ prime sur le maxi.

    C. Le tarif cote au conditionnement ET propose des paliers sur le NOMBRE
       de conditionnements. Ici seulement, si le prix est divisé par N, le
       palier est **multiplié par N** : 10 cartons de 100, ce sont 1 000
       unités, pas 10. Écrire le prix dégressif sur un palier 10 valoriserait
       à tort toute commande entre 10 et 999.

    Le test de tri : le tarif porte-t-il plus d'un prix pour la même référence ?
    Non, et il est au conditionnement -> B. Oui, dont un unitaire -> A.
    Oui, aucun unitaire -> C.

CE QUI SORT
    `diviseur_conditionnement`, `source_diviseur`, `prix_avant_division`,
    `preuve_diviseur`, `minimum_de_commande`. Sans ces colonnes, un prix faux
    est indétectable après coup.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

# --------------------------------------------------------------------------
# Rang 4 : ce qu'on accepte de déduire

# Conditionnements plausibles. On ne teste pas n'importe quel entier : un
# rapport de 7,0 est plus probablement une hausse qu'un lot de sept.
CONDITIONNEMENTS_PLAUSIBLES = (2, 3, 4, 5, 6, 8, 10, 12, 20, 24, 25, 30, 36,
                               48, 50, 100, 144, 200, 250, 500, 1000)

# Tolérance autour de l'entier. 5 % laisse passer une hausse tarifaire modérée
# par-dessus le conditionnement (x2,05 chez le Fournisseur G reste un x2)
# sans avaler un
# rapport franchement différent.
TOLERANCE = 0.06

# En deçà, une récurrence est une coïncidence, pas une règle de fournisseur.
ARTICLES_MINIMUM = 5

# Écart maximal toléré entre le rapport d'un article et le facteur établi pour
# son tarif. 1,5 laisse passer une hausse tarifaire par-dessus le
# conditionnement sans avaler un rapport d'un autre ordre de grandeur.
MARGE_FACTEUR = 1.5


# --------------------------------------------------------------------------
# Rang 2 : lire un conditionnement dans un libellé
#
# Chaque motif est ancré et exige un mot-clé ou une forme non ambiguë. Sans
# cela, « MODELE EXEMPLE 1 T3 » ou la référence « 1234567 » produiraient un
# conditionnement, et le prix serait divisé par un numéro de modèle.

MOTIFS_DESIGNATION = [
    # « 20X100ML », « 12X1L », « 6X500G » : N sous-unités de volume/masse.
    # C'est N qui est le conditionnement, pas le volume.
    (re.compile(r"\b(\d{1,4})\s*[xX*]\s*\d+(?:[.,]\d+)?\s*(?:ML|L|G|KG|CL)\b"),
     "N x volume"),
    # « 4 bidons de 5 L », « 12 flacons de 500 ml », « 6 boîtes de 100 ».
    # La forme est « N <contenant> DE <quantité> » : c'est N qui compte,
    # jamais la quantité qui suit.
    #
    # L'ORDRE COMPTE : ce motif passe AVANT celui à mot-clé ci-dessous.
    # Sur « 6 boîtes de 100 », le motif à mot-clé lirait « BOITES DE 100 »
    # et rendrait 100 — le contenu d'une boîte — là où le tarif vend six
    # boîtes. Le nombre qui PRÉCÈDE le contenant l'emporte sur celui qui
    # le suit. « BOITE DE 100 », sans nombre devant, tombe bien sur le
    # motif suivant.
    #
    # Relevé le 16/09 chez le Fournisseur I : « PRODUIT EXEMPLE 1 - 4
    # bidons de 5 L + 1 pompe » à 92,00 EUR contre un DPA de 21,30 —
    # 92,00/4 = 23,00, soit 1,08 fois le DPA (valeurs d'exemple).
    # Deux garde-fous, tirés d'un balayage des 5 082 libellés du périmètre,
    # où ce motif produisait deux faux positifs :
    #   - PAS précédé de « N° » : « DOIGTIER TAILLE M N°3 SACHET DE 100 »
    #     donnait 3, qui est une TAILLE ;
    #   - PAS précédé de « x » ou « X » : « COMPRESSE 10 X 10 CARTON DE »
    #     donnait 10, qui est une DIMENSION.
    (re.compile(r"(?<![\d.,°])(?<![xX]\s)"
                r"\b(\d{1,4})\s*(?:BIDONS?|FLACONS?|BOITES?|BOÎTES?|CARTONS?"
                r"|SACHETS?|POCHES?|TUBES?|PACKS?|BRIQUES?|BOUTEILLES?"
                r"|CUPS?|ROULEAUX?)\s+DE\b"),
     "N contenants de"),
    # « BTE/50 », « BTE DE 50 », « BOITE DE 100 », « CARTON DE 24 », « SACHET/4 ».
    # Le mot-clé fait toute la sûreté du motif : c'est lui qui distingue un
    # conditionnement d'un nombre quelconque.
    (re.compile(r"\b(?:BTE|BOITE|BOÎTE|CARTON|CT|SACHET|SACH|PAQUET|PQT|LOT|"
                r"POCHE|ETUI|DISTRIB(?:UTEUR)?)\s*(?:DE\s*|D[E']\s*|/\s*|\.\s*)?"
                r"(\d{1,4})\b"),
     "boîte / carton / lot"),
    # « PAR 100 », « PAR 50 »
    (re.compile(r"\bPAR\s+(\d{1,4})\b"), "« par N »"),
    # « - 5 PCS », « 42PCS », « 10 UNITES », « 16 GALETTES ».
    # Relevé le 16/09 chez le Fournisseur S — « DRAP DE REHAUSSEMENT
    # 140x110cm - 5 PCS » à 133,00 EUR contre un DPA de 30,20 : 133,00/5 =
    # 26,60, soit 0,88 fois le DPA (valeurs d'exemple). Le lot était écrit,
    # personne ne le lisait.
    #
    # « PIECES » ÉCRIT EN TOUTES LETTRES EST EXCLU, et ce n'est pas un
    # oubli. Sur les 5 082 libellés du périmètre, la forme longue ne
    # désignait JAMAIS une quantité : elle désigne le TYPE de produit —
    # « POCHE MODELE EXEMPLE 2 2 PIECES » est un système deux-pièces,
    # « COLLIER TRACHEO 2 PIECES » un collier en deux parties. Diviser
    # leur prix par deux était le seul effet possible de ce mot.
    # L'abréviation « PCS », elle, est toujours une quantité.
    # Le nombre ne doit être précédé NI d'un chiffre, NI d'une lettre, NI
    # d'un « ° » : « CHUT MODELE EXEMPLE 3 NOIR T38 UNITE » donnerait sinon 38,
    # qui est une POINTURE, et diviserait par 38 le prix d'un soulier.
    (re.compile(r"(?<![\d,°A-Z])\b(\d{1,4})\s*(?:PCS\b|UNITES?\b"
                r"|UNITÉS?\b|UVC\b|GALETTES?\b)"),
     "N pièces"),
    # « PACK X36 », « X24 PCS », « PROTÈGE SEAU X20 ».
    #
    # Motif le plus rentable et le plus dangereux. Deux garde-fous, tirés d'un
    # passage sur les 5 082 désignations du périmètre :
    #   - NON suivi d'une unité de mesure : « 130 X 90 CM » est un drap, pas un
    #     lot de 90 ; « 2 X 1 KG » est une masse ;
    #   - NON précédé d'un nombre : « 10 X 12 CM », « 70 X 110 » sont des
    #     dimensions, et la seconde valeur n'y est jamais un conditionnement.
    (re.compile(r"(?<![\d.,])\s[xX]\s?(\d{1,4})\b"
                r"(?!\s*(?:CM|MM|M\b|KG|G\b|ML|L\b|CL|PO|POUCE))"),
     "« xN »"),
    # Le motif « /N » nu a été RETIRÉ. Sur les désignations réelles il ne
    # ramenait que des tailles — « T: 54/56 » donnait 56, « 380/405 » donnait
    # 405, « T39/41 » donnait 41. Vingt-quatre faux positifs, aucun vrai :
    # diviser un prix par une taille est précisément l'erreur muette que ce
    # module existe pour empêcher. Les vraies formes « BTE/100 » et
    # « SACHET/50 » sont déjà prises par le motif à mot-clé ci-dessus.
]

# Un conditionnement lu hors de ces bornes est presque sûrement autre chose :
# une taille, une longueur, une année, une référence.
BORNES_DESIGNATION = (2, 5000)

# « LA PAIRE », « PAIRE DE … ». Le mot est reconnu, mais il ne PROPOSE
# jamais de diviser — il ne sert qu'à reconnaître un libellé du WMS déjà au
# lot, c'est-à-dire à EMPÊCHER une division. Voir
# diviseur_depuis_designation, qui porte la mesure qui a tranché.
MOTIF_PAIRE = re.compile(r"\b(?:LA\s+PAIRE|PAR\s+PAIRE|PAIRES?)\b")


def _conditionnement_brut(libelle) -> tuple[int | None, str]:
    """Le nombre d'unités qu'annonce un libellé, sans rien interpréter."""
    if libelle is None or (isinstance(libelle, float) and pd.isna(libelle)):
        return None, ""
    texte = str(libelle).upper()
    for motif, nom in MOTIFS_DESIGNATION:
        trouve = motif.search(texte)
        if not trouve:
            continue
        try:
            valeur = int(trouve.group(1))
        except (IndexError, ValueError):
            continue
        if BORNES_DESIGNATION[0] <= valeur <= BORNES_DESIGNATION[1]:
            return valeur, nom
    if MOTIF_PAIRE.search(texte):
        return 2, "paire"
    return None, ""


def diviseur_depuis_designation(designation_tarif,
                                designation_wms=None) -> tuple[int | None, str]:
    """Le conditionnement à DIVISER, lu dans les libellés.

    Le point délicat : un conditionnement lu dans un libellé ne dit pas qu'il
    faut diviser. Il dit ce que la ligne représente. Ce qui commande la
    division, c'est l'ÉCART entre ce que cote le tarif et ce que compte le
    WMS.

        tarif « BTE/50 », WMS à l'unité         -> diviser par 50
        tarif « LA PAIRE »                      -> ne JAMAIS diviser ici

    POURQUOI « PAIRE » NE DIVISE PLUS (16/09)
        Une paire de rampes, c'est UN produit : on ne vend pas une rampe
        seule. Le mot désigne le plus souvent la nature de l'article, pas
        son conditionnement — et parfois un simple composant, comme chez
        le Fournisseur AA où « 1 paire de bottes » fait partie d'un
        appareil de pressothérapie à 2 500 € (valeur d'exemple).

        La contre-épreuve par le libellé du WMS ne suffisait pas : elle
        exigeait le MOT « paire » des deux côtés, alors que le WMS dit le
        même chose au PLURIEL — « RAMPES TELESCOPIQUES 152CM ».

        Mesure du 16/09 sur la passe complète : ce rang a divisé
        5 articles, et les 5 étaient faux. Chez le Fournisseur H
        (REF-EXEMPLE-1) et le Fournisseur CS (REF-EXEMPLE-2), le tarif lu
        était au CENTIME égal au DPA du WMS — la preuve qu'il ne fallait
        pas diviser était dans la donnée.
        Aucun des 5 rapports au DPA ne corroborait un ×2.

        Et rien ne s'appuyait dessus : les 50 lignes du Fournisseur G, le cas pour
        lequel la règle avait été écrite, passent toutes par le rang 4,
        où la division est corroborée par le rapport au DPA. C'est là que
        « paire » doit se prouver, ligne à ligne, ou pas du tout.

    Le mot reste reconnu par `_conditionnement_brut` : sur le libellé du WMS
    il sert à dire « cet article est DÉJÀ au lot », donc à empêcher une
    division. 296 articles du périmètre en dépendent.

    Quand on ne dispose que du libellé du tarif, on ne conclut pas : c'est à
    l'appelant de fournir le libellé du WMS en regard.
    """
    cond_tarif, nom = _conditionnement_brut(designation_tarif)
    if not cond_tarif:
        return None, ""
    # « Paire » ne propose jamais de diviser. Voir ci-dessus : 5 divisions,
    # 5 fausses, et le seul fournisseur concerné passe par le rang 4.
    if nom == "paire":
        return None, ""
    cond_wms, _ = _conditionnement_brut(designation_wms)
    # Les deux côtés annoncent le même conditionnement : même base, rien à faire.
    if cond_wms and cond_wms == cond_tarif:
        return None, ""
    # Le WMS annonce un conditionnement DIFFÉRENT : le rapport des deux est
    # le seul diviseur défendable (tarif par 100, WMS par 10 -> diviser par 10).
    if cond_wms and cond_wms != cond_tarif:
        if cond_tarif % cond_wms == 0:
            return (cond_tarif // cond_wms,
                    f"désignation ({nom}, net du conditionnement du WMS)")
        return None, ""
    return cond_tarif, f"désignation ({nom})"


# --------------------------------------------------------------------------
# Rang 4 : déduction par le rapport, jamais seule


def entier_rond_le_plus_proche(rapport) -> int | None:
    """Le conditionnement plausible dont `rapport` est proche, s'il existe.

    Sert à ÉTABLIR le facteur d'un tarif, donc il est volontairement strict :
    on ne veut compter comme preuve que les rapports qui tombent juste.
    """
    if rapport is None or pd.isna(rapport) or rapport <= 0:
        return None
    for n in CONDITIONNEMENTS_PLAUSIBLES:
        if abs(rapport - n) <= TOLERANCE * n:
            return n
    return None


def suit_le_facteur(rapport, facteur) -> bool:
    """Cet article suit-il le conditionnement déjà établi pour son tarif ?

    Question différente de la précédente, et c'est tout l'enjeu. Une fois qu'on
    sait que le tarif cote au lot de N — établi sur des dizaines d'articles —
    il ne s'agit plus de redemander à chaque ligne de prouver le facteur. Il
    s'agit de trancher entre deux hypothèses : cette ligne est-elle au lot, ou
    déjà unitaire ?

    On tranche par la plus proche des deux en RAPPORT, pas en écart absolu :
    la frontière basse est la moyenne géométrique de 1 et de N. Pour N = 2,
    c'est 1,41 — un rapport de 1,86 est au lot, un rapport de 1,05 ne l'est
    pas. Sans cela, quatre chaussures du Fournisseur G dont le tarif avait
    bougé de 9 %
    restaient au prix de la paire.

    MAIS la moyenne géométrique seule est ruineuse sur les grands facteurs :
    pour N = 100 elle accepte tout rapport entre 10 et 1 000. Chez le
    Fournisseur AE, des prix à 377 fois le DPA se sont ainsi fait diviser
    par 100, et sont ressortis à 0,0050 EUR contre 0,50 EUR (valeurs
    d'exemple) — le « divisé deux fois » du
    brief, en pire. On ajoute donc une SECONDE condition, multiplicative et
    bornée : le rapport doit rester à moins de moitié en plus ou en moins
    du facteur. Pour N = 100, cela donne [66,7 ; 150] et non [10 ; 1000].

    Les deux conditions sont cumulatives. Un rapport qui n'y satisfait pas
    n'est pas divisé : il est signalé, et c'est tout.
    """
    if rapport is None or pd.isna(rapport) or rapport <= 0 or not facteur:
        return False
    n = float(facteur)
    plus_proche_du_lot = rapport > n ** 0.5
    dans_la_fourchette = (n / MARGE_FACTEUR) <= rapport <= (n * MARGE_FACTEUR)
    return plus_proche_du_lot and dans_la_fourchette


def deduire_par_systematicite(rapports: pd.Series) -> tuple[int | None, str, int]:
    """Le conditionnement que le fournisseur applique systématiquement.

    On ne regarde pas l'ampleur d'un écart mais sa RÉCURRENCE. Un même entier
    qui revient sur des dizaines d'articles n'est pas une hausse : c'est une
    unité de vente. Une vraie évolution tarifaire ne tombe jamais sur le même
    facteur pour autant de références.

    Retourne (diviseur, preuve, nombre d'articles concernés).
    """
    utiles = rapports.dropna()
    utiles = utiles[utiles > 0]
    if len(utiles) < ARTICLES_MINIMUM:
        return None, "", 0
    candidats = utiles.map(entier_rond_le_plus_proche).dropna()
    if candidats.empty:
        return None, "", 0
    comptes = candidats.value_counts()
    meilleur = int(comptes.index[0])
    combien = int(comptes.iloc[0])
    if combien < ARTICLES_MINIMUM:
        return None, "", combien
    # La récurrence doit dominer : si la moitié des articles seulement suit le
    # facteur, le tarif mélange les unités de vente et rien ne se déduit en
    # bloc. C'est le cas du Fournisseur BP, et il ne se force pas.
    part = combien / len(utiles)
    if part < 0.60:
        return None, (f"facteur {meilleur} sur {combien} des {len(utiles)} "
                      f"articles seulement — tarif mixte, non déductible"), combien
    preuve = (f"rapport au DPA proche de {meilleur} sur {combien} des "
              f"{len(utiles)} articles comparables ({part:.0%})")
    return meilleur, preuve, combien


# --------------------------------------------------------------------------
# Choix du diviseur, tous rangs confondus


def choisir_diviseur(conditionnement_tarif, conditionnement_supposé,
                     designation, conditionnement_wms,
                     deduit=None, preuve_deduite="") -> tuple[float, str, str]:
    """Le diviseur retenu, sa source et sa preuve — dans l'ordre de fiabilité.

    Retourne toujours un triplet exploitable : (1, "", "") quand rien n'est
    établi. Diviser par défaut serait pire que ne rien faire.
    """
    # Rang 1 — RETIRÉ, et il faut dire pourquoi pour qu'on ne le remette pas.
    #
    # Le brief demandait d'utiliser la colonne de conditionnement du tarif
    # (« Nbre unités par boîte » chez le Fournisseur O, « PCB » chez le
    # Fournisseur AL) comme
    # diviseur le plus fiable. C'est faux DANS CE PROJET, parce que le schéma
    # de extracteurs/base.py est explicite :
    #
    #     prix_achat_unitaire_ht   prix d'achat d'UNE unité, HT
    #     conditionnement          nombre d'unités par colis
    #     prix_colis_ht            prix_achat_unitaire_ht x conditionnement
    #
    # Le prix est DÉJÀ unitaire : les extracteurs ont fait la division. La
    # colonne `conditionnement` sert à remonter au prix du colis, pas à en
    # redescendre. C'est d'ailleurs pour cela que les Fournisseurs V, O, AL,
    # BM et T ressortaient déjà à une médiane de 1,000 au contrôle C2.
    #
    # Ce rang a coûté cher avant d'être vu : chez le Fournisseur AE,
    # dix-neuf prix ont été divisés une seconde fois et sont ressortis à
    # 0,0050 EUR contre un DPA de 0,50 EUR (valeurs d'exemple). La
    # couverture du périmètre en a perdu 91 articles.
    #
    # Le conditionnement du schéma alimente donc `minimum_de_commande`, qui
    # sert au réapprovisionnement, et rien d'autre.

    # Rang 2 — le libellé du tarif
    depuis_libelle, nom = diviseur_depuis_designation(designation)
    if depuis_libelle:
        return float(depuis_libelle), "désignation", nom

    # Rang 3 — la fiche article du WMS
    if (conditionnement_wms is not None
            and not pd.isna(conditionnement_wms)
            and float(conditionnement_wms) > 1):
        return (float(conditionnement_wms), "WMS",
                "conditionnement de la fiche article")

    # Rang 4 — la déduction, seulement si elle a été corroborée en amont
    if deduit:
        return float(deduit), "déduit", preuve_deduite

    return 1.0, "", ""


# --------------------------------------------------------------------------
# Application


def configuration(nb_prix_pour_la_reference: int, a_un_prix_unitaire: bool) -> str:
    """A, B ou C — le test de tri du brief, écrit une fois pour toutes."""
    if nb_prix_pour_la_reference <= 1:
        return "B"
    return "A" if a_un_prix_unitaire else "C"


def appliquer(table: pd.DataFrame,
              colonne_prix: str = "prix_achat_unitaire_ht",
              colonne_palier: str | None = "paliers") -> pd.DataFrame:
    """Divise les prix par leur conditionnement, et trace tout ce qu'elle fait.

    La table doit déjà porter `diviseur_conditionnement` et `source_diviseur`.
    La division et la multiplication du palier se font ICI, dans la même
    fonction et avec le même N : les séparer, c'est prendre le risque qu'un
    jour l'une soit appliquée sans l'autre.
    """
    t = table.copy()
    diviseur = pd.to_numeric(t["diviseur_conditionnement"],
                             errors="coerce").fillna(1.0)
    prix = pd.to_numeric(t[colonne_prix], errors="coerce")

    t["prix_avant_division"] = prix
    # Un diviseur INFÉRIEUR à 1 est une multiplication : le tarif cote à
    # l'unité quand le WMS compte au lot. Diviser par 1/10 revient à multiplier
    # par 10, et la même formule sert les deux sens.
    a_convertir = (diviseur > 1) | ((diviseur > 0) & (diviseur < 1))

    # Cas A : l'extracteur a déjà retenu une colonne unitaire, il n'y a rien à
    # diviser. On le dit en colonne plutôt que de laisser croire à un oubli.
    if "config_conditionnement" in t.columns:
        cas_a = t["config_conditionnement"] == "A"
        a_convertir = a_convertir & ~cas_a

    a_diviser = a_convertir
    t[colonne_prix] = prix.where(~a_convertir, prix / diviseur)

    # Cas C, et lui seul : les paliers sont exprimés en nombre de
    # conditionnements, donc ils se convertissent en unités avec le même N.
    if colonne_palier and colonne_palier in t.columns and "config_conditionnement" in t.columns:
        cas_c = (t["config_conditionnement"] == "C") & a_diviser
        if cas_c.any():
            palier = pd.to_numeric(t.loc[cas_c, colonne_palier], errors="coerce")
            t.loc[cas_c, colonne_palier] = palier * diviseur[cas_c]

    # Le conditionnement reste un minimum de commande : on ne peut pas acheter
    # 3 unités d'un produit vendu par 100. Cette contrainte sert au
    # réapprovisionnement, pas au palier, donc elle a sa propre colonne.
    t["minimum_de_commande"] = diviseur.where(diviseur > 1)
    return t


# --------------------------------------------------------------------------
# Intégration au rapprochement


def appliquer_au_rapprochement(fusion: pd.DataFrame,
                               colonne_prix="prix_achat_unitaire_ht",
                               colonne_dpa="pa_wms") -> pd.DataFrame:
    """Pose le diviseur sur une table de rapprochement, et divise ce qui doit l'être.

    Appelée depuis pa_completer.py, une fois le tarif rapproché et le DPA joint.
    Elle ajoute cinq colonnes et ne modifie le prix que lorsqu'elle peut dire
    POURQUOI. Sans ces colonnes, un prix faux est indétectable après coup.

    Prudence délibérée sur les paliers : quand le tarif porte des paliers, on
    ne sait pas d'ici s'ils comptent des unités (cas A, rien à diviser) ou des
    conditionnements (cas C, il faudrait aussi multiplier le palier). Les deux
    se ressemblent et se trompent en silence, alors on SIGNALE au lieu de
    trancher. Le cas C est rare ; s'y tromper ne l'est pas.
    """
    t = fusion.copy()
    n = len(t)
    vide = pd.Series([None] * n, index=t.index, dtype="object")
    t["diviseur_conditionnement"] = 1.0
    t["source_diviseur"] = ""
    t["preuve_diviseur"] = ""
    t["config_conditionnement"] = ""
    t["prix_avant_division"] = pd.to_numeric(t.get(colonne_prix), errors="coerce")

    if colonne_prix not in t.columns:
        return t

    cond_tarif = (t["conditionnement"] if "conditionnement" in t.columns else vide)
    suppose = (t["conditionnement_suppose"]
               if "conditionnement_suppose" in t.columns
               else pd.Series([False] * n, index=t.index))
    desi_tarif = (t["designation"] if "designation" in t.columns
                  else t.get("Désignation tarif", vide))
    desi_wms = t.get("Libellé déclinaison ^(1)", vide)

    # Le DPA du WMS est la seule contre-épreuve disponible ligne à ligne.
    # Sans lui, une division lue dans un libellé ne se vérifie devant rien.
    prix_lu = pd.to_numeric(t.get(colonne_prix), errors="coerce")
    dpa_lu = (pd.to_numeric(t[colonne_dpa], errors="coerce")
              if colonne_dpa in t.columns
              else pd.Series([float("nan")] * n, index=t.index))

    # --- rangs 1 à 3, ligne à ligne ---
    diviseurs, sources, preuves, ecartees = [], [], [], []
    for i in t.index:
        # L'extracteur a-t-il DÉJÀ ramené ce prix à l'unité ? C'est le cas dès
        # qu'il renseigne un conditionnement lu au tarif — chez le
        # Fournisseur O, la colonne « Nbre unités par boîte ». Redécouper
        # « Boîte de 30 » dans la désignation diviserait alors une seconde
        # fois : l'étui pénien REF-EXEMPLE-3 tombait de 1,900 EUR à
        # 0,0633 EUR (valeurs d'exemple), la même erreur que chez le
        # Fournisseur AE.
        deja_divise = False
        valeur_cond = cond_tarif.get(i)
        if (valeur_cond is not None and not pd.isna(valeur_cond)
                and float(valeur_cond) > 1 and not bool(suppose.get(i, False))):
            deja_divise = True

        d, s, p = choisir_diviseur(
            cond_tarif.get(i), suppose.get(i, False),
            None if deja_divise else desi_tarif.get(i), None)
        # Le libellé du WMS sert de contre-épreuve : s'il annonce le même
        # conditionnement que le tarif, les deux comptent la même chose.
        if s == "désignation":
            d2, s2 = diviseur_depuis_designation(desi_tarif.get(i),
                                                 desi_wms.get(i))
            if not d2:
                d, s, p = 1.0, "", ""
            else:
                d = float(d2)

        # SECONDE contre-épreuve, et c'est elle qui tranche : le RAPPORT AU DPA.
        #
        # Le libellé dit ce que la ligne représente ; il ne dit pas que le
        # WMS compte autrement. Mesure du 16/09 sur la passe complète : sur treize
        # divisions issues d'un libellé, ONZE partaient d'un tarif déjà égal au
        # DPA — souvent au centime — et le divisaient quand même.
        #
        #   Fourn. AF  « Téterelles MODELE EXEMPLE 5 XL (Lot de 2) »
        #            DPA 6,00  tarif 6,00  -> proposé 3,00. Le DPA du WMS
        #            était déjà celui du lot de deux.
        #   Fourn. K   « Rétroviseur (la paire) pour MODELE EXEMPLE 6 X4 »
        #            divisé par QUATRE : le « X4 » est le nom du modèle.
        #            DPA 15,00  tarif 15,00  -> proposé 3,75.
        #   Fourn. BK  « Séparateurs x6 Petits » : DPA 9,00 pour les six.
        #            (prix : valeurs d'exemple)
        #
        # Les deux divisions qui survivent sont celles que le rapport
        # corrobore : le Fournisseur J « 12X1L » part de 12,01 fois le DPA et
        # retombe à 1,001 ; le Fournisseur D « sachet de 4 » part de 2,67.
        #
        # Même juge qu'au rang 4, volontairement : une division se prouve de la
        # même façon quelle que soit la source qui l'a suggérée.
        ecartee = ""
        # Une correspondance posée à la main échappe à la contre-épreuve :
        # la personne a validé la ligne du tarif ET le prix attendu, ce qui
        # vaut mieux qu'un rapport. Voir pa_correspondances.
        validee = bool(t.at[i, "correspondance_validee"]) \
            if "correspondance_validee" in t.columns else False
        if s == "désignation" and d > 1 and not validee:
            rapport = None
            if pd.notna(dpa_lu.get(i)) and float(dpa_lu.get(i)) > 0 \
                    and pd.notna(prix_lu.get(i)):
                rapport = float(prix_lu.get(i)) / float(dpa_lu.get(i))
            if rapport is None:
                # Aucun DPA : rien ne corrobore, rien ne contredit. On divise
                # — le libellé reste une indication — mais on le DIT, comme au
                # rang 4, pour que la relecture sache où regarder.
                p = f"{p} — aucun DPA, division non corroborée"
            elif not suit_le_facteur(rapport, d):
                ecartee = (f"division par {d:g} écartée : le tarif est à "
                           f"{rapport:.3f}× le DPA, ce qui ne soutient pas "
                           f"un lot de {d:g} — lu « {p} »")
                d, s, p = 1.0, "", ""

        # Le cas inverse, et il est tout aussi coûteux : le tarif cote à
        # l'unité quand le WMS compte au LOT. « POCHES MODELE EXEMPLE 7
        # 1,5 L 10 U » a un DPA de 16,00 EUR pour dix poches (valeur
        # d'exemple) ; proposer 1,60 EUR le valoriserait dix fois trop bas.
        # On MULTIPLIE alors, au lieu de diviser.
        if d == 1.0 and deja_divise:
            lot_wms, nom_lot = _conditionnement_brut(desi_wms.get(i))
            if lot_wms and lot_wms > 1 and abs(lot_wms - float(valeur_cond)) < 0.01:
                d = 1.0 / float(lot_wms)
                s = "WMS au lot"
                p = (f"le WMS compte par {int(lot_wms)} ({nom_lot}) et le tarif "
                     f"à l'unité — prix REMULTIPLIÉ, pas divisé")

        diviseurs.append(d)
        sources.append(s)
        preuves.append(p)
        ecartees.append(ecartee)
    t["diviseur_conditionnement"] = diviseurs
    t["source_diviseur"] = sources
    # Une division refusée ne disparaît pas : elle se lit en colonne. Sans
    # cela, la ligne ressemblerait à une ligne qu'aucune règle n'a touchée,
    # et personne ne saurait qu'un conditionnement y a été vu puis écarté.
    t["division_ecartee"] = ecartees
    t["preuve_diviseur"] = preuves

    # --- rang 4 : la déduction, en bloc et seulement si elle est systématique ---
    prix = pd.to_numeric(t[colonne_prix], errors="coerce")
    dpa = pd.to_numeric(t.get(colonne_dpa), errors="coerce") if colonne_dpa in t.columns else None
    if dpa is not None:
        sans_source = t["source_diviseur"] == ""
        rapports = (prix / dpa).where(sans_source & (dpa > 0) & (prix > 0))
        deduit, preuve, combien = deduire_par_systematicite(rapports)
        if deduit:
            # Un article dont le libellé du WMS annonce DÉJÀ un lot n'entre pas
            # dans la déduction : son rapport s'explique autrement.
            deja_au_lot = desi_wms.map(lambda x: _conditionnement_brut(x)[0] or 0)
            eligible = sans_source & (deja_au_lot <= 1)

            # Le facteur est établi pour le TARIF, mais il ne s'applique pas
            # aveuglément à chaque ligne. Trois situations, trois traitements :
            #
            #   - le rapport de l'article colle au facteur -> on divise, et
            #     l'article corrobore lui-même la division ;
            #   - l'article n'a PAS de DPA -> aucune contre-épreuve possible,
            #     mais le tarif est établi au lot : on divise en le disant.
            #     Sans cela, les CHUT MODELE EXEMPLE 8 XTRA sortaient à
            #     60,00 € (valeur d'exemple) — le prix
            #     de la paire — au seul motif qu'ils n'ont jamais été achetés ;
            #   - le rapport est proche de 1 -> cet article est DÉJÀ unitaire au
            #     tarif. Le diviser ferait un prix deux fois trop bas. C'est le
            #     tarif mixte, et il se signale au lieu de se forcer.
            colle = rapports.map(lambda r: suit_le_facteur(r, deduit))
            sans_dpa = eligible & rapports.isna()
            deja_unitaire = eligible & rapports.notna() & ~colle

            vise = eligible & colle
            t.loc[vise, "diviseur_conditionnement"] = float(deduit)
            t.loc[vise, "source_diviseur"] = "déduit"
            t.loc[vise, "preuve_diviseur"] = preuve

            t.loc[sans_dpa, "diviseur_conditionnement"] = float(deduit)
            t.loc[sans_dpa, "source_diviseur"] = "déduit (sans DPA)"
            t.loc[sans_dpa, "preuve_diviseur"] = (
                preuve + " — article sans DPA : diviseur du tarif appliqué, "
                         "non corroboré individuellement")

            t.loc[deja_unitaire, "preuve_diviseur"] = (
                f"NON divisé : le fournisseur cote au lot de {deduit}, mais le "
                f"rapport au DPA de cet article ne le suit pas — ligne "
                f"probablement déjà unitaire au tarif")
            print(f"  conditionnement déduit /{deduit} — {int(vise.sum())} "
                  f"articles corroborés, {int(sans_dpa.sum())} sans DPA, "
                  f"{int(deja_unitaire.sum())} laissés tels quels")
        elif preuve:
            print(f"  conditionnement NON déduit : {preuve}")

    # --- configuration A / B / C ---
    a_des_paliers = pd.Series(False, index=t.index)
    for nom in ("paliers", "Paliers"):
        if nom in t.columns:
            a_des_paliers |= t[nom].notna() & (t[nom].astype(str).str.strip() != "")
    t["config_conditionnement"] = "B"
    t.loc[a_des_paliers, "config_conditionnement"] = "A ou C — à vérifier"

    # On ne divise que le cas B : un seul prix, au conditionnement.
    a_diviser = (pd.to_numeric(t["diviseur_conditionnement"],
                               errors="coerce").fillna(1) > 1)
    bloque = a_diviser & a_des_paliers
    if bloque.any():
        t.loc[bloque, "preuve_diviseur"] = (
            t.loc[bloque, "preuve_diviseur"]
            + " — NON APPLIQUÉ : le tarif porte des paliers, vérifier s'ils "
              "comptent des unités (cas A) ou des conditionnements (cas C)")
        print(f"  {int(bloque.sum())} diviseurs non appliqués : paliers au tarif, "
              f"cas A ou C à trancher")
        t.loc[bloque, "diviseur_conditionnement"] = 1.0

    t = appliquer(t, colonne_prix=colonne_prix, colonne_palier=None)
    applique = (pd.to_numeric(t["diviseur_conditionnement"],
                              errors="coerce").fillna(1) > 1)
    if applique.any():
        print(f"  {int(applique.sum())} prix ramenés à l'unité "
              f"(sources : {t.loc[applique, 'source_diviseur'].value_counts().to_dict()})")
    return t


# --------------------------------------------------------------------------
# Diagnostic — signaler, ne rien appliquer


def diagnostic(motif: str | None = None) -> pd.DataFrame:
    """Ce que le chantier conditionnement trouverait, fournisseur par fournisseur.

    Lit l'état par article — donc les prix DÉJÀ rapprochés, y compris ceux que
    le garde-fou x10 a refusés : ce sont précisément eux qu'un diviseur peut
    débloquer.
    """
    from extracteurs.base import chemin_lisible

    chemin = RACINE / "sortie" / "3-achats" / "pa_etat_par_article.xlsx"
    if not chemin.exists():
        print("pa_etat_par_article.xlsx absent : lancer run_pa.py d'abord")
        return pd.DataFrame()
    d = pd.read_excel(chemin_lisible(chemin), sheet_name=0, header=3)

    d["ancien"] = pd.to_numeric(d["Ancien PA"], errors="coerce")
    d["neuf"] = pd.to_numeric(d["Nouveau PA"], errors="coerce")
    d["ecarte"] = pd.to_numeric(d["PA écarté (à vérifier)"], errors="coerce")
    # Le prix du tarif, qu'il ait été retenu ou refusé par le garde-fou.
    d["tarif"] = d["neuf"].fillna(d["ecarte"])
    if motif:
        d = d[d["Fournisseur"].fillna("").str.upper().str.contains(motif.upper())]

    lignes = []
    for nom, sous in d.groupby("Fournisseur"):
        comparables = sous[(sous["ancien"] > 0) & (sous["tarif"] > 0)]
        if comparables.empty:
            continue
        rapports = comparables["tarif"] / comparables["ancien"]
        mediane = rapports.median()
        deduit, preuve, combien = deduire_par_systematicite(rapports)

        # Attention : pa_etat_par_article ne porte que la désignation du WMS,
        # pas celle du tarif. Ce qu'on lit ici dit donc ce que le WMS compte —
        # et un article du WMS déjà libellé « LA PAIRE » ou « BTE/50 » est
        # DÉJÀ au lot :
        # il ne faut surtout pas le diviser. La colonne sert à reconnaître ces
        # cas, jamais à proposer un diviseur.
        cote_wms = (sous["Désignation"].map(
            lambda x: _conditionnement_brut(x)[0]).dropna())
        libelle_dominant = (int(cote_wms.mode().iloc[0])
                            if not cote_wms.empty else None)
        part_au_lot = len(cote_wms) / len(sous) if len(sous) else 0

        lignes.append({
            "Fournisseur": nom,
            "Comparables": len(comparables),
            "Médiane": round(mediane, 3),
            "Hors fourchette": not (0.70 <= mediane <= 1.30),
            "Diviseur déduit": deduit,
            "Articles concernés": combien,
            "WMS déjà au lot": libelle_dominant,
            "Part au lot": round(part_au_lot, 2),
            "Médiane après division": (round(mediane / deduit, 3)
                                       if deduit else None),
            "Preuve": preuve,
        })

    res = pd.DataFrame(lignes)
    if res.empty:
        return res
    return res.sort_values(["Hors fourchette", "Comparables"], ascending=False)


def a_trancher() -> pd.DataFrame:
    """Les articles où le conditionnement demande un arbitrage humain.

    Trois signaux, tous tirés de ce qu'on a déjà sous la main :

      1. Le libellé du WMS dit « UNITE » ou « UNITAIRE » alors que le prix
         retenu est plusieurs fois le DPA. C'est le signal le plus sûr : chez
         le Fournisseur I, les 16 articles dont le tarif unitaire l'emporte
         sur le prix carton portent TOUS « UNITE » dans leur libellé WMS.
      2. Le libellé du WMS annonce un lot (« BTE/50 », « LOT DE 10 ») et le prix
         est très inférieur au DPA : le tarif cote alors à l'unité et c'est
         une MULTIPLICATION qu'il faudrait, pas une division.
      3. Le rapport au DPA tombe près d'un entier rond sans qu'aucune source
         n'ait justifié de diviser.

    On ne tranche rien : on présente.
    """
    from extracteurs.base import chemin_lisible

    chemin = RACINE / "sortie" / "3-achats" / "pa_etat_par_article.xlsx"
    if not chemin.exists():
        print("pa_etat_par_article.xlsx absent : lancer run_pa.py d'abord")
        return pd.DataFrame()
    d = pd.read_excel(chemin_lisible(chemin), sheet_name=0, header=3)

    d["ancien"] = pd.to_numeric(d["Ancien PA"], errors="coerce")
    d["neuf"] = pd.to_numeric(d["Nouveau PA"], errors="coerce")
    d["ecarte"] = pd.to_numeric(d["PA écarté (à vérifier)"], errors="coerce")
    d["tarif"] = d["neuf"].fillna(d["ecarte"])
    d = d[(d["ancien"] > 0) & (d["tarif"] > 0)].copy()
    d["rapport"] = d["tarif"] / d["ancien"]

    libelle = d["Désignation"].fillna("").str.upper()
    dit_unite = libelle.str.contains(r"\bUNITE\b|\bUNITAIRE\b", regex=True)
    lot_wms = d["Désignation"].map(lambda x: _conditionnement_brut(x)[0] or 0)

    lignes = []
    for i in d.index:
        r = d.at[i, "rapport"]
        signal = ""
        propose = None
        if dit_unite[i] and r >= 2:
            signal = ("le WMS dit « unité » mais le prix vaut "
                      f"{r:.1f}× le DPA — le tarif cote sans doute au lot")
            propose = entier_rond_le_plus_proche(r) or round(r)
        elif lot_wms[i] > 1 and r <= 0.7:
            signal = (f"le WMS compte par {int(lot_wms[i])} et le prix vaut "
                      f"{r:.2f}× le DPA — le tarif cote à l'unité, il faudrait "
                      f"MULTIPLIER")
            propose = -int(lot_wms[i])
        elif r >= 1.8:
            rond = entier_rond_le_plus_proche(r)
            if rond:
                signal = (f"rapport ×{r:.2f}, proche de {rond} — "
                          f"conditionnement possible, non confirmé")
                propose = rond
        if not signal:
            continue
        lignes.append({
            "Code article": d.at[i, "Code article"],
            "Désignation": d.at[i, "Désignation"],
            "Fournisseur": d.at[i, "Fournisseur"],
            "Ancien PA": round(d.at[i, "ancien"], 4),
            "Prix retenu": round(d.at[i, "tarif"], 4),
            "Rapport": round(r, 2),
            "Diviseur proposé": propose,
            "Prix si appliqué": (round(d.at[i, "tarif"] / propose, 4)
                                 if propose and propose > 0
                                 else (round(d.at[i, "tarif"] * -propose, 4)
                                       if propose else None)),
            "À trancher": signal,
            "Fichier source": d.at[i, "Fichier source PA"],
        })
    res = pd.DataFrame(lignes)
    if res.empty:
        return res
    # Le plus gros écart en premier : c'est là que l'erreur coûte le plus.
    return res.sort_values("Rapport", ascending=False)


def ecrire_a_trancher() -> Path | None:
    import mise_en_forme
    from main import SORTIE_ACHATS

    res = a_trancher()
    if res.empty:
        print("aucun cas à trancher")
        return None
    chemin = mise_en_forme.chemin_ecriture(
        SORTIE_ACHATS / "pa_conditionnement_a_trancher.xlsx")
    with pd.ExcelWriter(chemin, engine="openpyxl") as w:
        res.to_excel(w, sheet_name="À trancher", index=False, startrow=3)
    mise_en_forme.formater(chemin, {
        "À trancher": {
            "ligne_entete": 4, "figer_colonne": 3,
            "titre": "Conditionnement — les cas qui demandent un œil",
            "sous_titre": "Un diviseur négatif veut dire MULTIPLIER : le "
                          "WMS compte par lot, le tarif cote à l'unité"},
    })
    print(f"  {len(res)} cas à trancher -> {chemin}")
    return chemin


def main() -> None:
    if "--a-trancher" in sys.argv:
        ecrire_a_trancher()
        return
    motif = sys.argv[1] if len(sys.argv) > 1 else None
    res = diagnostic(motif)
    if res.empty:
        print("rien à diagnostiquer")
        return

    pd.set_option("display.width", 220)
    hors = res[res["Hors fourchette"]]
    print(f"{len(res)} fournisseurs comparables, "
          f"{len(hors)} hors de la fourchette 0,70-1,30")
    print()
    if not hors.empty:
        print("=== HORS FOURCHETTE — contrôle C2 ===")
        print(hors[["Fournisseur", "Comparables", "Médiane", "Diviseur déduit",
                    "Articles concernés", "Médiane après division",
                    "WMS déjà au lot", "Part au lot"]].to_string(index=False))
        print()

    proposables = res[res["Diviseur déduit"].notna()]
    if not proposables.empty:
        print("=== DIVISEURS CORROBORÉS PAR LA SYSTÉMATICITÉ ===")
        for _, r in proposables.iterrows():
            # Si le WMS compte DÉJÀ au lot, le rapport ne vient pas de là et la
            # division serait une erreur. On le dit au lieu de proposer.
            reserve = ""
            if r["WMS déjà au lot"] and r["Part au lot"] >= 0.5:
                reserve = (f"  [ATTENTION : le WMS est déjà au lot de "
                           f"{int(r['WMS déjà au lot'])} sur "
                           f"{r['Part au lot']:.0%} des articles — ne pas diviser "
                           f"sans vérifier]")
            print(f"  {r['Fournisseur'][:44]:<46} /{int(r['Diviseur déduit']):<4} "
                  f"{r['Médiane']:>7.3f} -> {r['Médiane après division']:<7.3f}{reserve}")
            print(f"      {r['Preuve']}")
    else:
        print("Aucun diviseur corroboré : rien ne s'applique en bloc.")

    print()
    print("Ce diagnostic SIGNALE. Rien n'est appliqué ici.")


if __name__ == "__main__":
    main()
