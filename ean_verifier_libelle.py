# -*- coding: utf-8 -*-
"""
Confronte au libelle les codes EUDAMED obtenus sur une reference courte.

    python ean_verifier_libelle.py             reprend ou il s'etait arrete
    python ean_verifier_libelle.py --perimetre le perimetre fige d'abord
    python ean_verifier_libelle.py --etat      ou en est-on ?
    python ean_verifier_libelle.py --essai 20  s'arrete apres 20 references

`--perimetre` ramene la collecte de 19 h a 8 h en n'interrogeant que les
articles du perimetre fige — ceux dont le code sert vraiment. Le cache
etant commun, elargir ensuite ne rejoue rien de ce qui est fait.

Sortie : wms_extracts/.cache_eudamed_libelle/resultats.json
         sortie/2-chantier-ean/eudamed_verification_libelle.xlsx

Le probleme
-----------
`eudamed_par_reference.py` interroge le registre sur le seul parametre
`reference`, SANS filtre fabricant, et retient le PREMIER dispositif
rendu dont le primaryDi est un EAN valide. Rien ne verifie que ce
dispositif est le notre.

Mesure du 14/09 : la reference « 7300 » (une ceinture 58 cm du
fournisseur BZ) rend 1 080 dispositifs, dont le premier est un tube
d'injection du fournisseur CA. La reference « 2975 » en rend 909, dont le
premier est une paire de lunettes de lecture italienne. Confrontes au
prefixe GS1 du tarif du
meme fournisseur, les codes ainsi obtenus ne concordent qu'a 27,8 %,
contre 100 % pour un code lu dans un tarif.

1 321 codes de V16 sont dans ce cas — 90,8 % de ce qu'EUDAMED a rendu.

Le libelle, seule caracteristique commune
-----------------------------------------
Nous n'avons pas le SRN des fabricants, et le registre n'accepte pas de
filtre sur leur nom (`manufacturerName` est ignore : la reponse reste
identique). En revanche `tradeName` est un VRAI filtre, en sous-chaine et
insensible a la casse, et il se combine avec `reference` cote serveur.

C'est ce qui change tout : l'absence devient demontrable. Demander
« reference=2950 ET tradeName=MODELE-B » interroge tout le registre et
rend zero — ce n'est plus une deduction tiree des cent premiers resultats
sur neuf cents, c'est un constat.

Et la meme requete repare autant qu'elle verifie. « reference=3000DBM ET
tradeName=DETACHABLE » rend deux dispositifs du fournisseur AH, dont
« Detachable Battery Module » — notre libelle mot pour mot, alors que le
code retenu jusqu'ici venait d'ailleurs.

Trois verdicts, jamais une correction d'office
-----------------------------------------------
    CONFIRME   un dispositif portant notre reference porte aussi notre
               libelle, et son code est celui qu'on avait deja
    CORRIGE    meme chose, mais le code differe : le notre etait faux et
               celui-ci est propose EN COLONNE, pas ecrit dans V16
    REJETE     aucun dispositif portant notre reference ne porte un
               libelle proche du notre, sur aucun de nos mots
    NON CONCLU le mot interroge n'etait pas assez distinctif (plus de
               cent candidats, aucun concordant) — a rejuger a la main
    NON TESTE  notre libelle n'offre aucun mot interrogeable

Ce script n'ecrit RIEN dans « EAN distributeur - Vn.xlsx », qui reste le
journal de ce que la chaine a produit. Il produit un avis ; c'est
`ean_fiable.py` et vous qui en tirez les consequences.

Choisir les mots a interroger
-----------------------------
Nos libelles sont francais, ceux du registre majoritairement anglais ou
allemands. « CEINTURE », « CHAUSSURES », « FAUTEUIL » n'y figureront
jamais : les interroger coute une minute pour rien. Ce sont les noms de
modele qui passent — MODELE-A, MODELE-B, M300 — et eux seuls.
D'ou `GENERIQUES`, a completer au fil des faux negatifs.

Mais une liste ecrite a la main ne couvrira jamais le vocabulaire
medical, et l'oubli se paie en minutes. Le tri se fait donc sur la
RARETE DANS NOS PROPRES LIBELLES, mesuree a chaque passe sur les
15 178 designations du classeur : un mot qu'on emploie partout ne
distingue rien, un mot qu'on emploie une fois est un nom de modele.

    GENOUILLERE ACTIVE MODELE-A S GA T5      MODELE-A   134 emplois
    CHAUSS. MODELE-B MARRON                  MODELE-B     1
    FAUTEUIL ROULANT FOURNISSEUR B M300 T42  M300        12

Trier sur la longueur, comme je l'avais fait d'abord, mettait
« GENOUILLERE » devant « MODELE-A » et brulait le premier appel. La
longueur ne sert plus que de departage a frequence egale.

Un mot de couleur ne suffit jamais. Un essai sur « CHAUSS. MODELE-B
MARRON » interroge sur MARRON a ramene une chaussure espagnole du
fournisseur GA, a 0,50 de score. C'est tout l'interet de garder le seuil
de `par_libelle` a 0,60 : ce rapprochement-la tombe de lui-meme.

Le throttling
-------------
Meme regle que la collecte par reference : un appel par minute. En
dessous, EUDAMED repond 200 avec zero resultat au lieu d'un 429, et un
echec devient indiscernable d'une absence — ce qui produirait exactement
le faux dont on essaie de sortir. L'etat est enregistre apres chaque
appel : une coupure coute au plus une requete.
"""

from __future__ import annotations

import collections
import json
import random
import re
import sys
import time
import unicodedata
from datetime import datetime
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import pandas as pd  # noqa: E402
import requests  # noqa: E402

import eudamed_fiabilite  # noqa: E402
import mise_en_forme  # noqa: E402
import par_libelle  # noqa: E402
import perimetre_liste  # noqa: E402
from extracteurs.base import chemin_lisible, ean_est_valide  # noqa: E402

URL = "https://ec.europa.eu/tools/eudamed/api/devices/udiDiData"
ENTETES = {
    "Accept": "application/json",
    "Accept-Language": "en",
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/120.0 Safari/537.36"),
}

SORTIE_CHANTIER = RACINE / "sortie" / "2-chantier-ean"
CACHE = RACINE / "wms_extracts" / ".cache_eudamed_libelle"
ETAT = CACHE / "resultats.json"
SORTIE = SORTIE_CHANTIER / "eudamed_verification_libelle.xlsx"

PAUSE = 60
JITTER = 12
TIMEOUT = 90

# Cent candidats par appel : c'est le maximum que le registre accepte de
# rendre d'un coup, verifie. Au-dela on paginerait, et une pagination a
# une minute la page coute plus qu'elle ne rapporte : si cent dispositifs
# portent notre reference ET un mot de notre libelle sans qu'aucun soit
# le notre, c'est le mot qui n'etait pas distinctif, pas le registre qui
# cache la reponse.
TAILLE_PAGE = 100

# Au-dela de trois mots, on interroge du vocabulaire courant et on paie
# une minute par mot. Les trois plus longs suffisent : ce sont les noms
# de modele.
MOTS_MAXIMUM = 3

# Le seuil de `par_libelle`, repris tel quel et volontairement non
# relache. A 0,50, un rapprochement sur la seule couleur passe.
SEUIL = par_libelle.SEUIL

# Variantes commerciales d'un meme dispositif. `par_libelle` mesure la
# part de NOS mots retrouvee chez l'autre, jamais les mots en trop — a
# raison, les catalogues decrivent plus longuement que nous. Mais un mot
# de cette liste present d'un seul cote ne decrit pas : il designe un
# autre article commercial.
#
# Le premier verdict reel l'a montre. « DETACHABLE BATTERY MODULE » a
# recu, a 1,00 de score, « Detachable Battery Module, RENTAL » — le
# module de location, pas celui qu'on vend.
#
# On ne rejette pas pour autant : c'est peut-etre le bon code, le
# registre n'ayant parfois qu'une declaration pour les deux. On le dit,
# en colonne, et c'est vous qui tranchez.
VARIANTES_COMMERCIALES = {
    "RENTAL", "LOCATION", "DEMO", "DEMONSTRATION", "SAMPLE", "TRIAL",
    "REFURBISHED", "RECONDITIONNE", "SPARE", "REPLACEMENT", "RECHANGE",
    "ACCESSORY", "ACCESSOIRE", "STERILE", "NONSTERILE", "DISPOSABLE",
    "REUSABLE", "SINGLEUSE",
}

# Vocabulaire francais qui n'a aucune chance de figurer dans un libelle
# du registre. A completer : chaque mot ajoute ici economise une minute
# par article qui le porte.
GENERIQUES = {
    "CHAUSS", "CHAUSSURE", "CHAUSSURES", "CEINTURE", "CEINTURES",
    "PROTECTEUR", "PROTECTION", "CUTANE", "CUTANEE", "STANDARD",
    "FAUTEUIL", "FAUTEUILS", "COUSSIN", "COUSSINS", "MATELAS",
    "SUPPORT", "SYSTEME", "MODULE", "TUBE", "TUBULURE", "POCHE", "POCHES",
    "SACHET", "SACHETS", "BOITE", "BOITES", "CARTON", "PAIRE", "PAIRES",
    "TAILLE", "POINTURE", "COLORIS", "MODELE", "GAMME", "SERIE",
    "DROIT", "DROITE", "GAUCHE", "AVANT", "ARRIERE", "AVEC", "SANS",
    "POUR", "PETIT", "PETITE", "GRAND", "GRANDE", "MOYEN", "MOYENNE",
    "UNIVERSEL", "UNIVERSELLE", "ADULTE", "ENFANT", "REGLABLE",
    "PLIANT", "PLIABLE", "FIXE", "MOBILE", "ELECTRIQUE", "MANUEL",
    "BLANC", "BLANCHE", "NOIR", "NOIRE", "BLEU", "BLEUE", "ROUGE",
    "VERT", "VERTE", "GRIS", "GRISE", "BEIGE", "TAUPE", "MARRON",
    "ROSE", "JAUNE", "ORANGE", "VIOLET", "ANTHRACITE", "CHOCOLAT",
}

COLONNES_SOURCE = ["Code article", "Référence", "Désignation", "Fournisseur",
                   "Réf. fournisseur", "Code EAN", "Source", "Fiabilité"]

# Fournisseurs qui nous ont envoye leur propre fichier de codes. On ne
# les interroge PAS : leur reponse tient a 98 % au test du prefixe GS1,
# EUDAMED a 27,8 %. Demander au registre ce que le fournisseur a deja
# ecrit, c'est payer une minute pour une reponse moins bonne — et risquer
# d'inscrire un code qui contredit la sienne.
#
# Le classeur consolide `ean_fournisseurs.xlsx` nomme ces reponses par
# FICHIER, pas par fournisseur : la correspondance ne peut pas s'en
# deduire, elle s'ecrit ici. A completer a chaque reponse recue.
#
# Le nom est celui du WMS, au caractere pres — releve dans V16.
REPONSE_RECUE = {
    "FOURNISSEUR B": "fichier EAN.pdf — 65 codes fauteuils",
    "FOURNISSEUR L": "fournisseur_l - demande codes EAN.xlsx",
    "LABORATOIRE FOURNISSEUR I": "Copie de fournisseur_i - demande codes EAN.xlsx",
    "FOURNISSEUR V": "Copie de fournisseur_v - demande codes EAN.xlsx",
    "FOURNISSEUR BM": "fournisseur_bm - reponse codes EAN.xlsx",
    "FOURNISSEUR BG FRANCE": "fournisseur_bg_-_demande_codes_EAN.xlsx",
    "FOURNISSEUR S S.A.": "FOURNISSEURS REFS ARTICLES ET CODE EAN 2025.xlsx",
    "FOURNISSEUR AU": "fournisseur_au - demande codes EAN.xlsx",
}

# Combien de nos libelles emploient chaque mot. Rempli par
# `apprendre_frequences`, lu par `mots_a_interroger`.
_FREQUENCES: dict[str, int] = {}


def _derniere_version() -> Path:
    """La version figee la plus recente.

    Trie sur le NUMERO, jamais sur le nom : l'ordre alphabetique place
    « V9 » apres « V16 », et `pa_completer.py` se fait encore prendre par
    la meme corde (son glob retient « EAN distributeur.xlsx », le fichier de
    travail).
    """
    fichiers = sorted(
        SORTIE_CHANTIER.glob("EAN distributeur - V*.xlsx"),
        key=lambda p: int(re.search(r"- V(\d+)", p.name).group(1)),
    )
    if not fichiers:
        raise SystemExit("Aucun fichier « EAN distributeur - Vn.xlsx ».")
    return fichiers[-1]


def _empecher_la_veille() -> bool:
    """Demande a Windows de ne pas s'endormir tant qu'on travaille.

    Reprise telle quelle de `eudamed_par_reference.py` : une collecte qui
    n'utilise ni clavier ni souris compte comme de l'inactivite, et le
    poste s'endort au bout de quelques dizaines de minutes. Sur huit
    heures, c'est la collecte entiere qu'on perd.

    On ne demande PAS ES_DISPLAY_REQUIRED : l'ecran peut s'eteindre.
    L'effet cesse avec le processus — rien a defaire.
    """
    if sys.platform != "win32":
        return False
    try:
        import ctypes
        ES_CONTINUOUS = 0x80000000
        ES_SYSTEM_REQUIRED = 0x00000001
        return bool(ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED))
    except Exception:
        # Une politique de domaine peut le refuser : ce n'est pas une
        # raison d'interrompre, seulement de le signaler.
        return False


def _charger() -> dict:
    if ETAT.exists():
        try:
            return json.loads(ETAT.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            pass
    return {}


def _enregistrer(etat: dict) -> None:
    CACHE.mkdir(parents=True, exist_ok=True)
    provisoire = ETAT.with_suffix(".tmp")
    provisoire.write_text(json.dumps(etat, ensure_ascii=False),
                          encoding="utf-8")
    provisoire.replace(ETAT)


def _sans_accent(texte) -> str:
    if texte is None or (isinstance(texte, float) and texte != texte):
        return ""
    decompose = unicodedata.normalize("NFKD", str(texte).upper())
    sans = "".join(c for c in decompose if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]+", " ", sans)


# Mots trop repandus dans les raisons sociales du secteur pour designer
# une marque. Sans ce filtre, « LABORATOIRE FOURNISSEUR I » et
# « COMPRESSE MEDICALE » partageraient un mot qui n'identifie rien. Meme liste que
# `pa_completer.MOTS_TROP_COURANTS`, recopiee plutot qu'importee : ce
# module ne doit pas dependre du chantier PA.
MOTS_TROP_COURANTS = {
    "CONFORT", "SANTE", "MEDICAL", "MEDICALE", "MEDICAUX", "FRANCE",
    "FRANCAISE", "SARL", "GROUPE", "LABORATOIRE", "LABORATOIRES",
    "INTERNATIONAL", "INTERNATIONALE", "DISTRIBUTION", "TECHNIQUE",
    "SERVICES", "EUROPE", "PHARMA", "SELF", "TOUS", "HYGIENE",
    "PRODUCT", "PRODUITS", "SOINS", "HEALTH", "CARE",
}

# Equivalences de taille entre notre notation et celle du registre.
#
# Le garde-fou des mots discriminants de `par_libelle` empeche qu'une
# taille L recoive le code d'une taille S — il est indispensable. Mais il
# compare des CHAINES, et « MEDIUM » ne ressemble pas a « M » :
#
#     MARQUE-U SLIP GAMME-1 MEDIUM      vs   MARQUE-U Slip Gamme-1 M 3x21p
#     MARQUE-U SLIP GAMME-2 EXTRA LARGE vs   MARQUE-U Slip Gamme-2 XL 3x21p
#
# Les deux rapprochements sont JUSTES et le garde-fou les rejetait, parce
# que {MEDIUM} et {M} different. On ramene donc les deux cotes a la meme
# notation avant de comparer. L'ordre compte : « EXTRA LARGE » doit etre
# vu avant « LARGE », sinon il devient « EXTRA L ».
TAILLES = [
    ("EXTRA EXTRA LARGE", "XXL"), ("EXTRA LARGE", "XL"),
    ("EXTRA SMALL", "XS"), ("X LARGE", "XL"), ("X SMALL", "XS"),
    ("TRES GRAND", "XL"), ("TRES PETIT", "XS"),
    ("MEDIUM", "M"), ("LARGE", "L"), ("SMALL", "S"),
    ("MOYEN", "M"), ("MOYENNE", "M"), ("GRAND", "L"), ("PETIT", "S"),
]


def harmoniser_tailles(texte) -> str:
    """Ramene « MEDIUM » et « M », « EXTRA LARGE » et « XL », a la meme forme."""
    normalise = f" {_sans_accent(texte)} "
    for longue, courte in TAILLES:
        normalise = normalise.replace(f" {longue} ", f" {courte} ")
    return normalise.strip()


def marque_probable(designation, fournisseur) -> str | None:
    """Le mot qui figure a la fois dans notre libelle et chez le fournisseur.

    C'est la marque, et c'est ce qu'EUDAMED indexe. « HYGIENE PRODUCT
    MARQUE-U » et « MARQUE-U DISCREET MAXI » partagent MARQUE-U : une
    requete `reference` + `tradeName=MARQUE-U` rend UN dispositif, le bon,
    la ou la reference seule en rendait des centaines.

    Le tri par rarete ne pouvait pas le trouver : MARQUE-U figure dans 88
    de nos libelles, donc il passait derriere GAMME-1 (9) ou CHANGE (2)
    et n'etait jamais soumis. La marque est frequente CHEZ NOUS et
    distinctive CHEZ EUX — les deux ne se contredisent pas.

    Sans intersection, on ne force rien : le premier mot d'un libelle
    n'est pas toujours la marque (« GENOUILLERE ACTIVE MODELE-A »), et
    en inventer une couterait une minute pour rien.
    """
    mots_fournisseur = {m for m in _sans_accent(fournisseur).split()
                        if len(m) >= 4 and m not in MOTS_TROP_COURANTS}
    if not mots_fournisseur:
        return None
    for mot in _sans_accent(designation).split():
        if mot in mots_fournisseur:
            return mot
    return None


def apprendre_frequences(designations) -> None:
    """Combien de nos libelles emploient chaque mot.

    Se calcule sur TOUS les libelles du classeur, pas sur les seuls
    articles a verifier : c'est notre vocabulaire d'ensemble qui dit
    qu'« ATTELLE » est courant et « MODELE-A » distinctif.
    """
    global _FREQUENCES
    compte: collections.Counter = collections.Counter()
    for designation in designations.dropna():
        compte.update({mot for mot in _sans_accent(designation).split()
                       if len(mot) >= 4 and not mot.isdigit()})
    _FREQUENCES = dict(compte)


def mots_a_interroger(designation, fournisseur=None) -> list[str]:
    """Les mots de notre libelle qui valent une requete, le plus rare d'abord.

    On ote le nom du fournisseur en tete — nos libelles le repetent, le
    registre non — puis tout ce qui est trop court, purement numerique ou
    trop courant pour distinguer quoi que ce soit.

    Si `apprendre_frequences` n'a pas ete appele, tous les mots pesent
    zero et l'ordre retombe sur la longueur : degrade, jamais faux.
    """
    texte = _sans_accent(designation)
    prefixe = _sans_accent(fournisseur).strip()
    if prefixe and texte.startswith(prefixe):
        texte = texte[len(prefixe):]

    retenus = []
    for mot in texte.split():
        if len(mot) < 4 or mot.isdigit():
            continue
        if mot in GENERIQUES or mot in par_libelle.MOTS_VIDES:
            continue
        if mot not in retenus:
            retenus.append(mot)
    retenus.sort(key=lambda mot: (_FREQUENCES.get(mot, 0), -len(mot)))

    # La marque passe DEVANT tout le reste : c'est la seule chose que
    # notre libellé et celui du registre nomment pareil.
    marque = marque_probable(designation, fournisseur)
    if marque:
        retenus = [marque] + [m for m in retenus if m != marque]
    return retenus[:MOTS_MAXIMUM]


def _code_unite(primary_di) -> str | None:
    """Le DI ramene a un EAN-13 d'unite de vente, ou rien.

    Un GTIN-14 d'indicateur 1 a 8 designe un CARTON, jamais l'unite : le
    retenir ferait entrer un code de regroupement dans un referentiel
    d'articles. Seul l'indicateur 0 est un EAN-13 prefixe d'un zero.
    """
    brut = str(primary_di or "").strip()
    if len(brut) == 14:
        if brut[0] != "0":
            return None
        candidat = brut[1:]
    elif len(brut) == 13:
        candidat = brut
    else:
        return None
    return candidat if ean_est_valide(candidat) else None


def _interroger(reference: str, mot: str) -> tuple[list[dict], int, str]:
    """(candidats, total annonce, statut) pour un couple reference/mot."""
    try:
        reponse = requests.get(
            URL, headers=ENTETES, timeout=TIMEOUT,
            params={"languageIso2Code": "en", "page": 0,
                    "size": TAILLE_PAGE, "pageSize": TAILLE_PAGE,
                    "reference": reference, "tradeName": mot})
    except requests.RequestException:
        return [], 0, "echec"

    if reponse.status_code != 200:
        return [], 0, "echec"
    if "json" not in reponse.headers.get("content-type", ""):
        return [], 0, "echec"

    donnees = reponse.json()
    total = donnees.get("totalElements", 0)
    # Un total demesure signale un filtre ignore, pas un resultat
    if total > 100_000:
        return [], 0, "echec"
    if not total:
        return [], 0, "absent"

    # Uniquement ce qui se serialise : le cache est du JSON, et les
    # ensembles de mots n'y entrent pas. Ils se reconstruisent a la
    # lecture, dans `_verdict` — c'est du calcul, pas de la donnee.
    candidats = []
    for element in donnees.get("content") or []:
        code = _code_unite(element.get("primaryDi"))
        if not code:
            continue
        candidats.append({
            "libelle": element.get("tradeName")
            or element.get("deviceName") or "",
            "ean": code,
            "reference": element.get("reference"),
            "fabricant": element.get("manufacturerName"),
        })
    return candidats, int(total), "trouve"


def _variante_commerciale(notre_libelle, libelle_eudamed) -> str:
    """Un mot de variante present du seul cote du registre, s'il y en a un."""
    les_notres = set(_sans_accent(notre_libelle).split())
    les_leurs = set(_sans_accent(libelle_eudamed).split())
    ecart = (les_leurs & VARIANTES_COMMERCIALES) - les_notres
    if not ecart:
        return ""
    return ("le libellé EUDAMED porte « " + ", ".join(sorted(ecart))
            + " », absent du nôtre : variante commerciale ?")


def _verdict(article: dict, etat: dict) -> dict:
    """Confronte un article au registre, mot par mot, et tranche.

    S'arrete au premier mot concluant : les suivants couteraient une
    minute pour confirmer ce qu'on sait deja.
    """
    fichier = REPONSE_RECUE.get(str(article["Fournisseur"] or "").strip())
    if fichier:
        return {"verdict": "RÉPONSE FOURNISSEUR",
                "motif": f"ne pas interroger EUDAMED : {fichier}"}

    reference = str(article["Réf. fournisseur"]).strip()
    mots = mots_a_interroger(article["Désignation"], article["Fournisseur"])
    if not mots:
        return {"verdict": "NON TESTÉ", "motif": "aucun mot interrogeable"}

    # Les mots reellement soumis. Un rejet ne vaut que par eux : « rejete
    # apres MODELE-B, CUTANE, OVALE » se lit, « rejete » tout court ne se
    # verifie pas. C'est la premiere chose qui manquait a l'essai du 14/09.
    essayes: list[str] = []
    vu_beaucoup = False
    for mot in mots:
        essayes.append(mot)
        cle = f"{reference}|{mot}"
        connu = etat.get(cle)
        if connu is None:
            candidats, total, statut = _interroger(reference, mot)
            connu = {"statut": statut, "total": total,
                     "candidats": candidats,
                     "date": datetime.now().strftime("%Y-%m-%d %H:%M")}
            etat[cle] = connu
            _enregistrer(etat)
            if statut != "echec":
                time.sleep(PAUSE + random.uniform(0, JITTER))

        if connu["statut"] == "echec":
            return {"verdict": "NON CONCLU", "motif": "appel en échec",
                    "mot": ", ".join(essayes)}
        if connu["statut"] == "absent":
            continue

        # Les ensembles de mots sont reconstruits ici plutot que
        # serialises — mais dans des COPIES. Les enrichir sur place
        # revenait a poser des `set` dans les dictionnaires du cache, qui
        # sont les memes objets : l'enregistrement suivant echouait, et
        # l'echec survenait une ligne APRES celle qui l'avait cause.
        # Les tailles sont harmonisees des DEUX cotes avant comparaison —
        # sinon « MEDIUM » et « M » se disqualifient l'un l'autre. Le
        # libelle d'origine reste intact pour l'affichage.
        candidats = []
        for candidat in (connu.get("candidats") or []):
            compare = harmoniser_tailles(candidat["libelle"])
            candidats.append({
                **candidat,
                "mots": par_libelle.mots(compare),
                "nombres": par_libelle.nombres(compare),
                "discriminants": par_libelle.discriminants(compare),
                "quantite": par_libelle.quantite(compare),
            })

        trouve = par_libelle.chercher(harmoniser_tailles(article["Désignation"]),
                                      candidats,
                                      marque=article["Fournisseur"],
                                      seuil=SEUIL)
        if trouve:
            actuel = str(article["Code EAN"] or "").strip()
            identique = trouve["ean"] == actuel
            return {
                "verdict": "CONFIRMÉ" if identique else "CORRIGÉ",
                "mot": mot,
                "score": trouve["score"],
                "libelle_eudamed": trouve["libelle"],
                "fabricant_eudamed": trouve["fabricant"],
                "ref_eudamed": trouve["reference"],
                "code_propose": trouve["ean"],
                "alerte": _variante_commerciale(article["Désignation"],
                                                trouve["libelle"]),
                "motif": ("le dispositif portant notre référence porte "
                          "aussi notre libellé"),
            }
        if connu.get("total", 0) > TAILLE_PAGE:
            vu_beaucoup = True

    if vu_beaucoup:
        return {"verdict": "NON CONCLU",
                "mot": ", ".join(essayes),
                "motif": (f"plus de {TAILLE_PAGE} dispositifs, aucun "
                          f"concordant : mot trop courant")}
    return {"verdict": "REJETÉ",
            "mot": ", ".join(essayes),
            "motif": ("aucun dispositif ne porte à la fois notre référence "
                      "et l'un de ces mots — le code actuel vient d'ailleurs")}


def _dans_le_perimetre(df: pd.DataFrame) -> pd.Series:
    """Les lignes appartenant au perimetre fige, qu'on LIT sans le recalculer.

    La cle est la Reference interne. Le code article ne sert qu'en repli,
    parce que quelques lignes du classeur EAN n'ont pas de reference — et
    perdre un article sur un champ vide serait le pire des filtres.
    """
    references = perimetre_liste.references()
    codes = perimetre_liste.codes_article()
    par_reference = df.get("Référence", pd.Series("", index=df.index))
    par_reference = par_reference.fillna("").str.strip().isin(references)
    par_code = df["Code article"].fillna("").str.strip().isin(codes)
    return par_reference | par_code


def _a_verifier(perimetre_seul: bool = False) -> pd.DataFrame:
    """Les lignes dont le code vient d'EUDAMED sur une reference courte."""
    fichier = _derniere_version()
    df = pd.read_excel(chemin_lisible(fichier), sheet_name="Codes EAN",
                       skiprows=3, dtype=str)
    # Le vocabulaire s'apprend sur TOUT le classeur, avant tout filtre :
    # c'est notre usage d'ensemble qui dit qu'un mot est courant.
    apprendre_frequences(df["Désignation"])
    df = df[df["Exploitable"].fillna("").str.strip() == "oui"]
    douteux = df[eudamed_fiabilite.douteux(df)].copy()
    print(f"{len(douteux)} codes à vérifier ({fichier.name})")

    if perimetre_seul:
        douteux = douteux[_dans_le_perimetre(douteux)].copy()
        print(f"  {perimetre_liste.entete()}")
        print(f"  {len(douteux)} dans le périmètre figé")

    print(f"  {douteux['Réf. fournisseur'].nunique()} références distinctes")
    print(f"  vocabulaire appris : {len(_FREQUENCES)} mots")
    return douteux


def _ecrire(lignes: list[dict]) -> None:
    if not lignes:
        return
    df = pd.DataFrame(lignes)
    ordre = ["CORRIGÉ", "REJETÉ", "NON CONCLU", "CONFIRMÉ",
             "RÉPONSE FOURNISSEUR", "NON TESTÉ"]
    df["_r"] = df["Verdict"].map({v: i for i, v in enumerate(ordre)})
    df = df.sort_values(["_r", "Fournisseur", "Désignation"]).drop(
        columns="_r")

    synthese = (df["Verdict"].value_counts()
                .rename_axis("Verdict").reset_index(name="Codes"))

    SORTIE_CHANTIER.mkdir(parents=True, exist_ok=True)
    chemin = mise_en_forme.chemin_ecriture(SORTIE)
    with pd.ExcelWriter(chemin, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Verdicts", index=False, startrow=3)
        synthese.to_excel(writer, sheet_name="Synthèse", index=False,
                          startrow=3)

    corriges = int((df["Verdict"] == "CORRIGÉ").sum())
    rejetes = int((df["Verdict"] == "REJETÉ").sum())
    confirmes = int((df["Verdict"] == "CONFIRMÉ").sum())
    mise_en_forme.formater(chemin, options={
        "Verdicts": {
            "ligne_entete": 4,
            "titre": "Codes EUDAMED confrontés au libellé",
            "sous_titre": (
                f"{len(df)} codes examinés | {confirmes} confirmés, "
                f"{corriges} à corriger, {rejetes} à retirer | "
                f"aucune écriture dans « EAN distributeur - Vn.xlsx »"),
        },
        "Synthèse": {
            "ligne_entete": 4,
            "titre": "Répartition des verdicts",
            "sous_titre": ("un code n'est retenu que si un dispositif porte "
                           "À LA FOIS notre référence et notre libellé"),
        },
    })
    print(f"\n→ {chemin.name}")


def etat_courant(perimetre_seul: bool = False) -> None:
    etat = _charger()
    articles = _a_verifier(perimetre_seul)
    refs = set(articles["Réf. fournisseur"].astype(str).str.strip())
    interrogees = {cle.split("|", 1)[0] for cle in etat}
    print(f"  {len(etat)} appels en cache")
    print(f"  {len(refs & interrogees)} références déjà entamées sur "
          f"{len(refs)}")
    restantes = len(refs - interrogees)
    print(f"  reste au moins {restantes} références, soit environ "
          f"{restantes * PAUSE / 3600:.0f} h au rythme d'un appel/minute")


def main() -> None:
    perimetre_seul = "--perimetre" in sys.argv
    if "--etat" in sys.argv:
        etat_courant(perimetre_seul)
        return

    limite = None
    if "--essai" in sys.argv:
        position = sys.argv.index("--essai")
        limite = int(sys.argv[position + 1]) if position + 1 < len(sys.argv) \
            else 10

    if not limite:
        if _empecher_la_veille():
            print("veille système désactivée le temps de la collecte "
                  "(l'écran peut s'éteindre)")
        else:
            print("! la mise en veille n'a pas pu être empêchée : une nuit "
                  "de collecte peut être perdue si le poste s'endort")

    etat = _charger()
    articles = _a_verifier(perimetre_seul)
    if limite:
        articles = articles.head(limite)
        print(f"  essai : {len(articles)} articles seulement")

    lignes = []
    for rang, article in enumerate(articles.to_dict("records"), start=1):
        resultat = _verdict(article, etat)
        lignes.append({
            "Code article": article["Code article"],
            "Référence": article.get("Référence"),
            "Désignation": article["Désignation"],
            "Fournisseur": article["Fournisseur"],
            "Réf. fournisseur": article["Réf. fournisseur"],
            "Code EAN actuel": article["Code EAN"],
            "Verdict": resultat["verdict"],
            "Code EAN proposé": resultat.get("code_propose"),
            "Libellé EUDAMED": resultat.get("libelle_eudamed"),
            "Fabricant EUDAMED": resultat.get("fabricant_eudamed"),
            "Réf. EUDAMED": resultat.get("ref_eudamed"),
            "Mot interrogé": resultat.get("mot"),
            "Concordance": resultat.get("score"),
            "Alerte": resultat.get("alerte"),
            "Motif": resultat["motif"],
        })
        print(f"  [{rang:>5}/{len(articles)}] {resultat['verdict']:<11} "
              f"{str(article['Réf. fournisseur'])[:14]:<14} "
              f"{str(article['Désignation'])[:42]:<42} "
              f"{resultat.get('code_propose') or ''}", flush=True)

        if rang % 25 == 0:
            _ecrire(lignes)

    _ecrire(lignes)


if __name__ == "__main__":
    main()
