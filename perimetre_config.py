# -*- coding: utf-8 -*-
"""
Perimetre articles — parametres figes.

CE FICHIER PORTE DES DECISIONS METIER, PAS DES REGLAGES
=======================================================
Les valeurs ci-dessous ont ete validees. Elles ne sont pas des
parametres a optimiser, et personne ne les modifie de sa propre
initiative — ni un script, ni un agent, ni une « amelioration ».

Un chiffre qu'on ajuste pour ameliorer un resultat cesse d'etre un
critere : il devient une variable d'ajustement, et le classement qu'il
produit ne veut plus rien dire. Si l'un d'eux doit changer, c'est une
decision explicite, elle se date, et le numero de version ci-dessous
change avec elle.

Toute execution qui lit ce fichier doit reporter PERIMETRE_VERSION dans
sa sortie : sans quoi on ne sait pas, devant un classement, sur quelles
regles il a ete produit.
"""

from __future__ import annotations

# =====================================================================
# DECISIONS METIER VALIDEES — NE PAS MODIFIER
# =====================================================================

# Fenetre d'observation des ventes, en mois.
# C'est la duree sur laquelle on regarde si un article a tourne.
FENETRE_VENTE_MOIS = 12

# Nombre de ventes minimum, sur la fenetre, pour qu'un article soit
# considere comme tournant. Une seule vente suffit.
SEUIL_TOURNE = 1

# A quelle frequence le classement se rejoue.
CADENCE_RECALCUL = "trimestrielle"

# Regle d'entree au perimetre (validee le 03/09/2026).
#
#   PERIMETRE = pas de motif ecarte
#               ET (vente sur 12 mois OU achat recu sur 12 mois)
#
# L'achat recu est un critere d'ENTREE, pas seulement un signal de
# confirmation : un article approvisionne est un article vivant, meme
# s'il ne s'est pas encore vendu. C'est le seul moyen de rattraper les
# articles neufs et ceux dont la vente n'est pas encore remontee.
#
# L'ORDRE COMPTE : les motifs d'ecartement s'appliquent EN PREMIER. Un
# article ecarte (SE, SV, FO, 19, 79, LO, N903, ancienne nomenclature)
# n'entre pas, meme avec un achat recent.
ACHAT_VAUT_ENTREE = True

# Troisieme porte d'entree, validee le 08/09/2026 : le CATALOGUE PRO.
#
#   Le perimetre retient ce qui BOUGE. Le catalogue professionnel
#   retient ce qu'on PUBLIE. Les deux ne coincident pas — et un article
#   publie sans prix d'achat reste un article publie sans prix d'achat,
#   quel que soit son historique de mouvement.
#
# L'ORDRE NE CHANGE PAS : les motifs d'ecartement s'appliquent toujours
# EN PREMIER. Une reference du catalogue ecartee par un motif (SE, SV,
# FO, 19, 79, LO, N903, ancienne nomenclature) n'entre PAS — 58 sont
# dans ce cas, dont 29 « produit arrete fabricant » qui portent l'arret
# dans leur propre libelle. Elles sont documentees comme « ecartee ET au
# catalogue pro » : un article arrete qui reste publie est un signal a
# porter au metier, pas une ligne a traiter en silence.
CATALOGUE_PRO_VAUT_ENTREE = True

# PERIMETRE FIGE. La regle ne bouge que sur decision explicite, et
# aucun article n'est reclasse de sa propre initiative.
#
# Volume attendu : 5 082 articles. Toute execution qui rend un autre
# nombre, a extraits inchanges, signale une regression — pas une
# evolution.
#   v1.1  4 731  vente OU achat recu
#   v1.2  5 082  + 351 references du catalogue pro sans activite
PERIMETRE_FIGE = True
VOLUME_ATTENDU = 5082

# Les articles sans vente ni achat sortent avec le motif SANS_ACTIVITE.
# Ils RESTENT au referentiel — ils ne sont simplement pas travailles.
# Sortir du perimetre n'est pas sortir de la base.
MOTIF_SANS_ACTIVITE = "SANS_ACTIVITE"

# Portee des chantiers, decidee avec le perimetre :
#   PA  : ne traiter que les articles du perimetre. Les ecarts hors
#         perimetre sont documentes et clos, pas repris.
#   EAN : EUDAMED tourne sur TOUT le referentiel — le cache est
#         constitue, rejouer ne coute rien, donc rien ne justifie de
#         restreindre une interrogation gratuite.
CHANTIERS_LIMITES_AU_PERIMETRE = ("PA",)

# EUDAMED n'est pas borne au perimetre : une requete automatisee sur un
# cache deja constitue ne coute ni argent ni temps humain. La restriction
# au perimetre ne vaut que pour ce qui coute quelque chose — un mail a un
# fournisseur, un releve en entrepot.
EUDAMED_SUR_TOUT_LE_REFERENTIEL = True

# ---------------------------------------------------------------------
# Fin des decisions metier. Ce qui suit en decoule.
# ---------------------------------------------------------------------

# Date de validation de ces valeurs.
DATE_VALIDATION = "2026-09-03"

# Version du perimetre. A reporter dans toute sortie produite avec ces
# regles, et a incrementer si l'une d'elles change.
#   1.0  fenetre 12 mois, seuil 1 vente, recalcul trimestriel
#   1.1  l'achat recu sur 12 mois devient un critere d'entree
#   1.2  le catalogue pro devient une troisieme porte d'entree
PERIMETRE_VERSION = "1.2"

# Traduction de la cadence en mois, pour le calcul de la prochaine
# echeance. Une seule source de verite : la cadence reste l'ecrit
# ci-dessus, ceci n'en est que la lecture machine.
CADENCES_EN_MOIS = {
    "mensuelle": 1,
    "trimestrielle": 3,
    "semestrielle": 6,
    "annuelle": 12,
}
CADENCE_RECALCUL_MOIS = CADENCES_EN_MOIS[CADENCE_RECALCUL]


def resume() -> str:
    """Les regles en vigueur, en une ligne, pour l'en-tete d'une sortie."""
    entree = "vente OU achat reçu" if ACHAT_VAUT_ENTREE else "vente"
    if CATALOGUE_PRO_VAUT_ENTREE:
        entree += " OU catalogue pro"
    return (f"périmètre v{PERIMETRE_VERSION} — "
            f"fenêtre {FENETRE_VENTE_MOIS} mois, "
            f"seuil {SEUIL_TOURNE} vente, "
            f"entrée : {entree}, "
            f"recalcul {CADENCE_RECALCUL} — "
            f"validé le {DATE_VALIDATION}")


if __name__ == "__main__":
    print(resume())
