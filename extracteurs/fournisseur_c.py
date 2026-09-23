# -*- coding: utf-8 -*-
"""
Extracteur du tarif du FOURNISSEUR C (PDF, « tarif_fournisseur.pdf »,
25 pages).

Appareils de PPC/ventilation, masques, consommables et accessoires. Le PDF a
une couche texte propre — aucun OCR n'est necessaire.

UNE SEULE COLONNE DE PRIX, ET C'EST LE PRIX D'ACHAT
    Le document est le tarif nominatif adresse au distributeur. Il ne publie
    ni prix public, ni prix conseille, ni LPPR : une seule colonne de
    montants, intitulee « Tarif 2026 HT en Euros » (pages masques et
    appareils) ou « HT en € » (pages accessoires, en paysage).

    Verification, faite avant d'ecrire la premiere regex : sur les 1 567
    lignes de la couche texte, UNE SEULE porte deux montants, et c'est un
    abonnement de service (« 1,00€/semaine soit 4,33€/mois », valeurs
    d'exemple). Il n'existe donc aucune seconde colonne de prix a confondre
    avec la premiere — le piege habituel (public HT / public TTC / remise)
    ne se pose pas ici.

    Recoupement avec les ordres de grandeur connus du marche : un appareil
    MODELE EXEMPLE a 600,00 EUR (tarif public constate autour de 900 EUR),
    un masque MODELE EXEMPLE a 100,00 EUR (public ~ 160 EUR). Ce sont bien
    des prix d'achat distributeur, pas des prix publics. (montants :
    valeurs d'exemple)

POURQUOI LA GEOMETRIE ET NON UNE REGEX SUR LE TEXTE
    Le texte brut de pypdfium2 colle la designation et le prix, et le tarif
    contient des designations qui FINISSENT par un nombre :

        K10000  Kit d'alimentation MODELE EXEMPLE 10  500,00 EUR

    Lu en texte, « EXEMPLE 10 500,00 » donne 10 500,00 EUR pour un kit
    d'alimentation a 500 EUR — un facteur 21 qui passerait le filet du ×10 de
    pa_completer une fois sur deux. Les abscisses tranchent sans ambiguite :
    le « 10 » est a x=204, dans la colonne designation, et le prix commence a
    x=443. On lit donc le PDF avec pdfplumber, par colonnes.

    Deuxieme raison : sur les pages accessoires, pdfplumber coupe les
    montants n'importe ou — « 24,00 » sort en deux mots « 2 » puis « 4,00 »,
    « 413,00 » en « 4 » puis « 13,00 ». Seule la concatenation de TOUT ce qui
    se trouve dans la colonne de droite reconstitue le nombre.

TROIS MISES EN PAGE DANS LE MEME FICHIER
    masques (portrait)     : Reference | (A fuite) | Designation | Prix | COMMENTAIRES
    appareils (portrait)   : Reference | Designation | Prix
    accessoires (paysage)  : Reference | Designation | 13 colonnes de
                             compatibilite machine | Prix

    Les bornes ne sont donc pas codees en dur : la colonne de prix est
    reperee page par page a partir de l'abscisse des symboles « € », qui sont
    tous alignes. Ce qui se trouve a DROITE du « € » n'est pas un prix mais
    la colonne COMMENTAIRES (« ARRET DE COMMERCIALISATION 31 MAI 2025 ») :
    elle part en alerte, jamais en designation.

    Le symbole n'est pourtant pas toujours la : la page DIVERS n'en porte
    aucun, et trois blocs de gammes l'omettent ligne a ligne. La colonne se
    cale alors sur le bord droit des montants — sans quoi une page entiere
    disparaissait, dont la contrainte mentonniere REF0001, qui est au
    perimetre.

    Les croix de compatibilite machine (« x » isolees, entre la designation
    et le prix) sont retirees : ce ne sont pas des mots de la designation.

LES LIGNES SE LISENT SUR PLUSIEURS RANGS
    Une ligne sur cinq a sa designation sur un rang different de sa
    reference et de son prix. Trois figures, toutes presentes :

        rang 220 : REF0002  x x x x  2 2,00 EUR    prix avant designation
        rang 221 :        Filtre MODELE EXEMPLE, Std, pack de 12

        rang 385 : Systeme de polysomnographie ambulatoire MODELE EXEMPLE
        rang 391 : REF0003-KA                        15 000,00 EUR
        rang 397 : chargeur, piles et kit de demarrage Adultes

    D'ou le regroupement par proximite verticale (TOLERANCE_RANG) plutot
    qu'un decoupage en lignes de texte. Le controle d'exhaustivite est
    simple et il tombe juste : 551 cellules de prix dans le PDF, 551 lignes
    lues.

CE QU'ON NE DIVISE PAS
    Beaucoup de references sont des lots : « Filtre MODELE EXEMPLE (par 12) »
    a 15,00 EUR. Ce ne sont pas des conditionnements a diviser mais des
    REFERENCES DISTINCTES — le meme filtre existe a l'unite (REF0004,
    1,50 EUR), par 2 (REF0005), par 12 (REF0006), par 50 (REF0007), chacune
    avec son propre code que le WMS stocke tel quel. Diviser donnerait un
    prix qu'aucune commande ne pourrait passer. Le lot est donc signale en
    alerte et le prix reste celui de la reference. (references et montants :
    valeurs d'exemple)

LES REPETITIONS NE SONT PAS DES DOUBLONS
    Le meme accessoire est reimprime dans chaque gamme compatible : le clip
    magnetique REF0008 revient dans huit rubriques, le gabarit REF0009 dans
    six. 61 des 551 lignes lues sont de cette nature, toujours au meme prix —
    aucune reference du tarif ne porte deux prix differents, verifie. Elles
    sont ramenees a une ligne par (reference, prix).

CE QUI EST VOLONTAIREMENT LAISSE DE COTE
    - Les 5 dernieres pages de produits sont intitulees « References en fin
      de commercialisation » : aucune colonne de prix, uniquement
      « Jusqu'a epuisement des stocks ». Rien a extraire.
    - Les montants a 0,00 EUR (gabarits de mesure offerts, logiciel fourni
      a « - € ») : un PA a zero injecte dans le WMS est plus dangereux qu'un
      PA absent. Ils sont comptes, pas retournes.
    - Les abonnements de service sont cotes « 1€/semaine soit 4,33€/mois » :
      ce n'est pas un prix unitaire d'article.
"""

from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path

import pandas as pd

from extracteurs.base import chemin_lisible, finaliser, nettoyer_texte

FOURNISSEUR = "FOURNISSEUR C"

# Ecart vertical maximal, en points PDF, entre deux fragments d'une meme
# ligne de tarif. Le regroupement se fait de proche en proche : quand la
# designation tient sur deux niveaux, la reference et le prix sont cales
# ENTRE les deux, et la ligne se lit donc dans cet ordre
#
#     rang 385   Systeme de polysomnographie ambulatoire MODELE EXEMPLE
#     rang 391   REF0003-KA                                 15 000,00 EUR
#     rang 397   chargeur, piles et kit de demarrage Adultes
#
# soit des sauts de 6 pt au plus, quand deux lignes de tarif voisines sont
# toujours a 7 pt au moins (11 pt sur les pages appareils). 6 pt separent
# donc les deux cas sans jamais les confondre.
TOLERANCE_RANG = 6.0

# Abscisse de fin de la colonne « References ». Elle demarre a x=36 sur les
# pages portrait comme sur les pages paysage ; la designation ne commence
# jamais avant x=80. La frontiere separe aussi les titres de rubrique
# (« MASQUES FACIAUX », cales a gauche) des lignes d'article.
X_FIN_REFERENCES = 70.0

# Distance verticale au-dela de laquelle une mention de la colonne
# COMMENTAIRES n'est plus rattachee a une ligne : elle porte alors sur toute
# une rubrique, pas sur un article.
ECART_COMMENTAIRE = 8.0

# Abscisse a partir de laquelle commencent les colonnes de compatibilite
# machine des pages accessoires (une colonne par gamme compatible), cochees
# par un « x » minuscule isole. La premiere est a x=243, la designation la
# plus longue s'arrete a x=231 : 230 separe les deux sans les confondre.
#
# La condition d'abscisse n'est pas un luxe. Un filtre pose sur le seul
# caractere effacait le « X » de « Coussins narinaires - X Small » (REF0010),
# qui devenait le jumeau exact du Small (REF0011) — deux articles distincts
# rendus indiscernables par le libelle.
X_DEBUT_COMPATIBILITE = 230.0

# Largeur maximale d'une cellule de prix, en points. Sert a remonter depuis
# le « € » jusqu'au premier chiffre du montant : « 16 000,00 » (valeur d'exemple) occupe
# 38 pt, on prend 45 pt de marge. Au-dela commencerait la colonne precedente.
LARGEUR_CELLULE_PRIX = 45.0

# Une reference du FOURNISSEUR C : 12345, 12345-KB, 10000001, R100-200,
# K10000, 27000K, 7000001-PK12, 7000002-3 (valeurs d'exemple). Au moins
# trois chiffres — c'est ce qui distingue une reference d'un titre de
# rubrique commencant par un chiffre (« 2 NIVEAUX DE PRESSION ») ou d'un
# abonnement (« ABOEXEMPLE01 »).
REFERENCE = re.compile(r"^[A-Z]{0,4}\d[A-Z0-9]*(?:[-/][A-Z0-9]+)*$")

# Montant reconstitue apres recollage des morceaux de la colonne de droite.
MONTANT = re.compile(r"^\d{1,6},\d{2}$")

# Rangs a ignorer : en-tetes repetes en haut de chaque page et pied de page.
BRUIT = re.compile(
    r"Internal Use|R[ée]f[ée]rences\s+D[ée]signation|COMMENTAIRES"
    r"|TARIFS? (MASQUES|APPAREILS|ACCESSOIRES)|en Euros|Tarif 2026 HT"
    r"|fin de commercialisation",
    re.IGNORECASE,
)

# Lots vendus sous une reference propre : « (par 12) », « pack de 50 »,
# « (x12) », « (25) ». Sert uniquement a poser une alerte — voir le docstring.
LOT = re.compile(
    r"\(\s*par\s+(\d+)\s*\)|pack\s+de\s+(\d+)|\(\s*x\s*(\d+)\s*\)|\(\s*(\d{2,})\s*\)",
    re.IGNORECASE,
)


def _mot_texte(mot) -> str:
    return mot["text"].strip()


def _rangs(mots: list[dict]) -> list[list[dict]]:
    """Regroupe les mots d'une page en rangs visuels.

    pdfplumber rend les mots dans le desordre : on trie par ordonnee puis on
    chaine tant que l'ecart reste sous TOLERANCE_RANG.
    """
    rangs: list[list[dict]] = []
    courant: list[dict] = []
    precedente = None
    for mot in sorted(mots, key=lambda m: (m["top"], m["x0"])):
        # Le chainage se fait sur l'ordonnee du mot PRECEDENT, pas sur celle
        # du premier mot du rang : une ligne a deux niveaux progresse par
        # sauts de 4 pt et s'etale sur 8 pt au total.
        if precedente is not None and mot["top"] - precedente > TOLERANCE_RANG:
            rangs.append(courant)
            courant = []
        courant.append(mot)
        precedente = mot["top"]
    if courant:
        rangs.append(courant)
    return rangs


def _colonne_prix(mots: list[dict]) -> tuple[float, float] | None:
    """Bornes (gauche, droite) de la colonne de prix de la page.

    Les symboles « € » sont tous alignes sur la meme abscisse ; le montant
    est cale a leur gauche, le commentaire eventuel a leur droite. On prend
    leur MEDIANE et non leur minimum : l'en-tete des pages accessoires
    contient lui aussi un « € », 12 points plus a gauche que la colonne, et
    il decalerait la frontiere jusque dans les croix de compatibilite.

    Toutes les pages ne portent pas le symbole : la page DIVERS n'en a aucun
    et trois blocs de gammes l'omettent ligne a ligne.
    A defaut, la colonne est calee sur le bord DROIT des montants, qui est
    exactement la ou se tiendrait le symbole. Sans ce recours, une page
    entiere disparaissait — dont la contrainte mentonniere REF0001, qui est
    au perimetre.

    Renvoie None sur une page sans le moindre montant (« References en fin
    de commercialisation », conditions generales de vente).
    """
    abscisses = sorted(m["x0"] for m in mots if _mot_texte(m) == "€")
    if not abscisses:
        bords = sorted(m["x1"] for m in mots if MONTANT.match(_mot_texte(m)))
        if not bords:
            return None
        # Le symbole manquant se poserait deux points apres le montant.
        abscisses = [bords[len(bords) // 2] + 2.0]
    mediane = abscisses[len(abscisses) // 2]
    return mediane - LARGEUR_CELLULE_PRIX, mediane + 10.0


def _intitule(page_texte: str) -> str:
    """Intitule EXACT de la colonne de prix, tel qu'il est imprime.

    Deux mises en page, deux libelles : les pages accessoires (paysage)
    abregent en « HT en € » la ou les pages masques et appareils ecrivent
    « Tarif 2026 HT en Euros ».
    """
    if "TARIF ACCESSOIRES APPAREILS 2026" in page_texte:
        return "HT en €"
    return "Tarif 2026 HT en Euros"


def _decouper(rang: list[dict], bornes: tuple[float, float]) -> dict:
    """Range les mots d'un rang dans les trois zones de la grille."""
    gauche, _ = bornes
    zone = {"reference": [], "designation": [], "prix": []}
    # Tri par ordonnee PUIS abscisse : une designation sur deux niveaux se
    # lit de haut en bas, pas de gauche a droite ; ses deux morceaux
    # commencent a la meme abscisse.
    for mot in sorted(rang, key=lambda m: (m["top"], m["x0"])):
        texte = _mot_texte(mot)
        if not texte:
            continue
        if mot["x0"] >= gauche:
            if texte != "€":
                zone["prix"].append(texte)
        elif mot["x0"] < X_FIN_REFERENCES:
            zone["reference"].append(texte)
        elif not (texte == "x" and mot["x0"] >= X_DEBUT_COMPATIBILITE):
            # Un « x » minuscule isole, passe la designation, est une croix
            # de compatibilite machine et non un mot du libelle.
            zone["designation"].append(texte)
    return zone


def _commentaires(mots: list[dict], droite: float) -> list[tuple[float, str]]:
    """Colonne COMMENTAIRES des pages masques, par ordonnee.

    Elle est a DROITE du prix et n'est pas alignee sur les lignes : une
    mention « ARRET DE COMMERCIALISATION 31 MAI 2025 » se pose a mi-hauteur
    entre deux articles, a 4 points de chacun. La laisser dans le
    regroupement des rangs souderait les deux lignes voisines en une seule
    et ferait disparaitre un article sur deux. Elle est donc mise de cote
    ici, puis rattachee a la ligne la plus proche.
    """
    groupes: dict[float, list[dict]] = {}
    for mot in mots:
        if mot["x0"] < droite or not _mot_texte(mot):
            continue
        clef = next((c for c in groupes if abs(c - mot["top"]) <= 2), mot["top"])
        groupes.setdefault(clef, []).append(mot)
    return [
        (clef, " ".join(_mot_texte(m) for m in sorted(v, key=lambda m: m["x0"])))
        for clef, v in sorted(groupes.items())
    ]


def _montant(morceaux: list[str]) -> float | None:
    """Recolle les morceaux de la colonne de droite en un montant.

    Sur les pages accessoires, « 24,00 » arrive en deux mots (« 2 », « 4,00 »)
    et « 413,00 » en « 4 » + « 13,00 » : on concatene sans espace, sinon le
    montant est ampute de son chiffre de tete.
    """
    brut = "".join(morceaux).replace(" ", "").replace(" ", "").replace(" ", "")
    if not MONTANT.match(brut):
        return None
    return float(brut.replace(",", "."))


def _lot(designation: str) -> int | None:
    trouve = LOT.search(designation or "")
    if not trouve:
        return None
    for valeur in trouve.groups():
        if valeur:
            quantite = int(valeur)
            # « (25x10mm) » et autres cotes ne sont pas des lots.
            return quantite if 1 < quantite <= 500 else None
    return None


def extract(path) -> pd.DataFrame:
    chemin = Path(path)
    try:
        import pdfplumber
    except ImportError:  # pragma: no cover
        return finaliser([])

    brutes: list[dict] = []
    sans_prix = 0
    a_zero = 0
    # Dernier titre de rubrique rencontre. Une poignee de lignes ont leur
    # cellule designation vide : le tarif se contente alors du titre juste
    # au-dessus (REF0012 sous « Tuyau avec bague de fuite MODELE EXEMPLE »).
    rubrique: str | None = None

    with pdfplumber.open(str(chemin_lisible(chemin))) as pdf:
        for page in pdf.pages:
            mots = page.extract_words()
            bornes = _colonne_prix(mots)
            if bornes is None:
                continue
            intitule = _intitule(page.extract_text() or "")
            grille = [m for m in mots if m["x0"] < bornes[1]]

            lignes_page: list[dict] = []
            courante: dict | None = None
            for rang in _rangs(grille):
                texte_rang = " ".join(_mot_texte(m) for m in rang)
                if BRUIT.search(texte_rang):
                    # Un en-tete ferme la ligne en cours, il ne l'annule pas :
                    # la derniere ligne d'une page se trouve juste au-dessus
                    # du pied de page et serait perdue sans cette ecriture.
                    if courante:
                        lignes_page.append(courante)
                    courante = None
                    continue

                zone = _decouper(rang, bornes)
                reference = next(
                    (r for r in zone["reference"] if REFERENCE.match(r) and
                     sum(c.isdigit() for c in r) >= 3),
                    None,
                )

                # Un titre de rubrique est cale a gauche, dans la colonne des
                # references, mais il DEBORDE sur la colonne designation
                # (« Kit Bulle MODELE EXEMPLE F40 » court de x=37 a x=96). Le
                # reconnaitre a sa seule abscisse de depart evite de coller
                # « F40 » a la fin de la designation de la ligne precedente.
                debut = min(m["x0"] for m in rang)
                if reference is None and debut < X_FIN_REFERENCES:
                    rubrique = nettoyer_texte(texte_rang)
                    if courante:
                        lignes_page.append(courante)
                    courante = None
                    continue

                if reference is not None:
                    if courante:
                        lignes_page.append(courante)
                    courante = {
                        "reference": reference,
                        "designation": list(zone["designation"]),
                        "prix": list(zone["prix"]),
                        "commentaire": [],
                        "colonne_prix": intitule,
                        "rubrique": rubrique,
                        # Ordonnee de la ligne, pour y raccrocher son
                        # commentaire une fois la page entiere lue.
                        "y": min(m["top"] for m in rang),
                    }
                    continue

                if courante is None:
                    continue
                # Rang sans reference : suite de la ligne precedente.
                courante["designation"] += zone["designation"]
                courante["prix"] += zone["prix"]

            if courante:
                lignes_page.append(courante)

            for hauteur, texte in _commentaires(mots, bornes[1]):
                if BRUIT.search(texte) or not lignes_page:
                    continue
                proche = min(lignes_page, key=lambda l: abs(l["y"] - hauteur))
                if abs(proche["y"] - hauteur) <= ECART_COMMENTAIRE:
                    proche["commentaire"].append(texte)

            brutes += lignes_page

    lignes = []
    for brute in brutes:
        prix = _montant(brute["prix"])
        designation = nettoyer_texte(" ".join(brute["designation"]))
        if prix is None:
            sans_prix += 1
            continue
        if prix == 0:
            # Gabarits offerts, logiciel a « - € » : un PA nul injecte dans
            # le WMS serait pris pour un vrai prix.
            a_zero += 1
            continue

        alertes = []
        if not designation:
            designation = brute.get("rubrique")
            alertes.append("designation absente de la ligne, reprise du titre "
                           "de rubrique")
        commentaire = nettoyer_texte(" ".join(brute["commentaire"]))
        if commentaire:
            alertes.append(commentaire)
        quantite = _lot(designation)
        if quantite:
            alertes.append(
                f"reference vendue par {quantite} — prix NON divise, "
                f"c'est le prix de la reference telle qu'elle se commande"
            )

        lignes.append({
            "fournisseur": FOURNISSEUR,
            "ref_fournisseur": brute["reference"],
            "designation": designation,
            "prix_achat_unitaire_ht": prix,
            "colonne_prix": brute["colonne_prix"],
            "fichier_source": chemin.name,
            "alerte": " | ".join(alertes) or None,
        })

    # Le tarif repete une meme reference dans chaque gamme compatible : le
    # clip REF0008 revient dans huit rubriques. On ne garde qu'une ligne par
    # (reference, prix) — et si une reference porte deux prix differents,
    # les deux sortent, signalees, plutot qu'un arbitrage silencieux.
    par_reference = defaultdict(set)
    for ligne in lignes:
        par_reference[ligne["ref_fournisseur"]].add(ligne["prix_achat_unitaire_ht"])

    vues = set()
    retenues = []
    for ligne in lignes:
        clef = (ligne["ref_fournisseur"], ligne["prix_achat_unitaire_ht"])
        if clef in vues:
            continue
        vues.add(clef)
        prix_multiples = par_reference[ligne["ref_fournisseur"]]
        if len(prix_multiples) > 1:
            liste = " / ".join(f"{p:.2f}" for p in sorted(prix_multiples))
            complement = f"reference a plusieurs prix dans le tarif : {liste}"
            ligne["alerte"] = (
                f"{ligne['alerte']} | {complement}" if ligne["alerte"] else complement
            )
        retenues.append(ligne)

    print(f"    {len(retenues)} lignes retenues sur {len(lignes)} lues "
          f"({len(lignes) - len(retenues)} repetitions d'une gamme a l'autre), "
          f"{sans_prix} sans prix, {a_zero} a 0,00 EUR")
    return finaliser(retenues)


if __name__ == "__main__":
    import sys

    defaut = (Path(__file__).resolve().parent.parent / "catalogues"
              / "fournisseur_c" / "2026" / "tarif_fournisseur.pdf")
    cible = Path(sys.argv[1]) if len(sys.argv) > 1 else defaut
    table = extract(cible)
    pd.set_option("display.width", 220)
    pd.set_option("display.max_colwidth", 70)
    pd.set_option("display.max_rows", 400)
    print(f"{len(table)} lignes extraites de {cible.name}")
    if not table.empty:
        prix = table["prix_achat_unitaire_ht"]
        print(f"prix : min {prix.min():.2f} / median {prix.median():.2f} / "
              f"max {prix.max():.2f} EUR")
        colonnes = ["ref_fournisseur", "designation", "prix_achat_unitaire_ht",
                    "colonne_prix"]
        print(table[colonnes].head(60).to_string(index=False))
        print("...")
        print(table[colonnes].tail(40).to_string(index=False))
