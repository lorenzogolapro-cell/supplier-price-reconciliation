# -*- coding: utf-8 -*-
"""
Accuses de reception du FOURNISSEUR D.

Expediteur commandes@fournisseur-d.example, piece jointe
Order_Acknowledgement.pdf. Le meme dossier Outlook recoit AUSSI les
avis d'expedition (expeditions@), qui ne portent aucun prix : le filtre
d'extraction se fait sur l'adresse complete, pas sur le domaine.

CE QUE CET AR A PROUVE (21/09/2026)
    La remise de 25 % (taux d'exemple) du fournisseur D etait jusque-la
    MESUREE dans les prix que le responsable des prix avait arbitres
    (quotient 0,7500, ecart-type 0,0001 sur 40 articles). Solide, mais
    circulaire : s'il s'etait trompe, la mesure aurait reproduit son
    erreur sans jamais la contredire — c'est exactement ce qui s'est
    produit sur les pages « TARIFS NETS », ou un double escompte etait
    valide par un DPA portant la meme erreur.

    Cet AR casse le cercle : il porte le taux EN CLAIR, colonne « Taux de
    remise % », a 25,00 sur 113 des 144 lignes lues.

    Il prouve aussi, separement, que les CHUT sont cotees a la PAIRE :
    rapprochee au tarif, une ligne de CHUT donne un rapport de 1/1,50 la
    ou un article normal donne 1/0,75. L'AR facture « 1 Each » ce que
    le WMS compte au soulier.

LES EXCEPTIONS, QUI VALENT LA LECTURE
    Une regle vraie a 90 % cache une seconde regle dans les 10 % restants.
      - une famille de chaussures (ref. 10-0001-x) est facturee a un
        autre taux, et le PU ne bouge pas entre les quantites 1, 2 et 3 :
        ce n'est pas du volume. Pose en `taux_par_famille` dans
        pa_remises.py.
      - une famille d'attelles est facturee « remise 0,00 % », ecrit noir
        sur blanc.
      - la ceinture d'electrostimulation 10-0002 affiche « 25,00 % » et
        est pourtant facturee au prix brut du tarif : sur cette famille,
        la colonne remise est decorative. Ne jamais lire un taux sans
        verifier le PU qui va avec.

LE PRIX RETENU
    « Prix Unitaire » est le prix NET — la remise y est deja appliquee,
    et le total de ligne vaut PU x quantite. Le taux affiche sert donc a
    comprendre, pas a recalculer.
"""

from __future__ import annotations

import re

import pandas as pd

from ar.base import cadrer, nettoyer_texte, nombre, premier_groupe, texte_pdf

FOURNISSEUR = "FOURNISSEUR D"

# DEUX MISES EN PAGE, ET LA PLUS ANCIENNE N'A PAS « Each »
#
#   2026 : 310U-47 MODELE-F BLUE EURO 47/... 1 Each FR 18-SEP-26
#          TRANSPORTEUR-Parcel-Suivi 13 40,00 25,00 40,00 EUR 5,50
#   2025 : 10-0003-3 MODELE-G H7.5CM SIZE 3 1 FR 22-SEP-25
#          TRANSPORTEUR-Parcel-Suivi 13 10,00 25,01 10,00 EUR 5,50
#
# Exiger « Each » rendait muets 37 des 70 AR — soit toute l'annee 2025 —
# sans que rien ne le signale : un AR sans ligne reconnue ressemble a un
# AR sans article. L'ancrage se fait donc sur ce que les deux formats ont
# en commun : quantite, code pays, date promise.
#
# Les trois nombres qui comptent sont les derniers avant « EUR ». Ce qui
# les separe du libelle (transporteur, numero de suivi) varie d'un AR a
# l'autre et ne se modelise pas : on l'absorbe.
LIGNE = re.compile(
    r"^(?P<ref>[A-Z0-9][A-Z0-9./+-]{2,24})\s+(?P<lib>.+?)\s+"
    r"(?P<qte>\d{1,4})\s+(?:Each\s+)?[A-Z]{2}\s+\d{2}-[A-Z]{3}-\d{2}\s+.*?"
    r"(?P<pu>\d{1,3}(?:[ .]\d{3})*,\d{2})\s+"
    r"(?P<remise>\d{1,2},\d{2})\s+"
    r"(?P<total>\d{1,3}(?:[ .]\d{3})*,\d{2})\s+EUR")
NUM_AR = re.compile(r"num.ro\s+(\d{5,})")
COMMANDE = re.compile(r"Votre\s+(?:r.f.rence|commande)\s*:?\s*(\d{6,})")
DATE = re.compile(r"Date de Commande:\s*(\d{2}-[A-Z]{3}-\d{2})")

MOIS = {"JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
        "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12}


def _date(texte: str):
    """« 18-SEP-26 » -> Timestamp. Ce fournisseur date en anglais abrege."""
    trouve = DATE.search(texte)
    if not trouve:
        return None
    jour, mois, annee = trouve.group(1).split("-")
    if mois not in MOIS:
        return None
    return pd.Timestamp(2000 + int(annee), MOIS[mois], int(jour))


def extract(chemin) -> pd.DataFrame:
    texte = texte_pdf(chemin)
    num_ar = premier_groupe(NUM_AR, texte)
    commande = premier_groupe(COMMANDE, texte)
    date_commande = _date(texte)

    lignes = []
    for brute in texte.splitlines():
        trouve = LIGNE.match(brute.strip())
        if not trouve:
            continue
        # Chez ce fournisseur le « Prix Unitaire » est deja net : le taux
        # affiche documente la remise, il ne reste pas a l'appliquer.
        unitaire = nombre(trouve.group("pu"))
        lignes.append({
            "num_ar": num_ar,
            "commande_wms": commande,
            "ref_fournisseur": nettoyer_texte(trouve.group("ref")),
            "designation_ar": nettoyer_texte(trouve.group("lib")),
            "quantite": nombre(trouve.group("qte")),
            "prix_unitaire_ar": unitaire,
            "remise_pct": nombre(trouve.group("remise")),
            "prix_net_ar": unitaire,
            "montant_ligne": nombre(trouve.group("total")),
            "date_commande": date_commande,
        })
    return cadrer(lignes, FOURNISSEUR, chemin)
