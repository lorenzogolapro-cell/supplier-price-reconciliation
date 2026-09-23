# -*- coding: utf-8 -*-
"""
Order acknowledgements from SUPPLIER D.

Sender commandes@fournisseur-d.example, attachment
Order_Acknowledgement.pdf. The same Outlook folder ALSO receives the
shipping notices (expeditions@), which carry no price at all: the
extraction filter is on the full address, not on the domain.

WHAT THIS ACKNOWLEDGEMENT PROVED (21/09/2026)
    Supplier D's 25 % discount (example rate) had until then been
    MEASURED in the prices the price owner had arbitrated (ratio 0.7500,
    standard deviation 0.0001 over 40 articles). Solid, but circular: had
    he been wrong, the measurement would have reproduced his error
    without ever contradicting it, which is exactly what happened on the
    "TARIFS NETS" pages, where a double discount was validated by a
    purchase price carrying the same error.

    This acknowledgement breaks the circle: it carries the rate IN PLAIN
    SIGHT, in the "Taux de remise %" column, at 25.00 on 113 of the 144
    rows read.

    It also proves, separately, that CHUT (temporary therapeutic
    footwear) are quoted by the PAIR: matched against the price list, a
    CHUT row gives a ratio of 1/1.50 where a normal article gives
    1/0.75. The acknowledgement invoices as "1 Each" what the WMS counts
    by the shoe.

THE EXCEPTIONS, WHICH ARE WORTH THE READING
    A rule that is true 90 % of the time hides a second rule in the
    remaining 10 %.
      - one shoe family (ref. 10-0001-x) is invoiced at a different
        rate, and the unit price does not move between quantities 1, 2
        and 3: this is not volume. Recorded as `taux_par_famille` in
        pa_remises.py.
      - one splint family is invoiced at "remise 0,00 %", written in
        black and white.
      - the 10-0002 electrostimulation belt shows "25,00 %" and is
        nevertheless invoiced at the price list's gross price: on that
        family, the discount column is decorative. Never read a rate
        without checking the unit price that goes with it.

THE PRICE KEPT
    "Prix Unitaire" is the NET price: the discount is already applied,
    and the line total equals unit price x quantity. The rate displayed
    is therefore there to be understood, not to be recomputed.
"""

from __future__ import annotations

import re

import pandas as pd

from ar.base import cadrer, nettoyer_texte, nombre, premier_groupe, texte_pdf

FOURNISSEUR = "FOURNISSEUR D"

# TWO LAYOUTS, AND THE OLDER ONE HAS NO "Each"
#
#   2026 : 310U-47 MODELE-F BLUE EURO 47/... 1 Each FR 18-SEP-26
#          TRANSPORTEUR-Parcel-Suivi 13 40,00 25,00 40,00 EUR 5,50
#   2025 : 10-0003-3 MODELE-G H7.5CM SIZE 3 1 FR 22-SEP-25
#          TRANSPORTEUR-Parcel-Suivi 13 10,00 25,01 10,00 EUR 5,50
#
# Requiring "Each" made 37 of the 70 acknowledgements silent, the whole
# of 2025, with nothing to signal it: an acknowledgement with no
# recognised row looks like an acknowledgement with no article. The
# anchor is therefore what the two formats have in common: quantity,
# country code, promised date.
#
# The three numbers that matter are the last ones before "EUR". What
# separates them from the label (carrier, tracking number) varies from
# one acknowledgement to the next and cannot be modelled: we absorb it.
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
    """Reads "18-SEP-26" as a Timestamp: this supplier dates in
    abbreviated English."""
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
        # At this supplier the "Prix Unitaire" is already net: the rate
        # displayed documents the discount, there is nothing left to
        # apply.
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
