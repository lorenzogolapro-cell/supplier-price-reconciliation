# -*- coding: utf-8 -*-
"""
Chantier PA — rapprocher par NOM DE MODÈLE, pour les tarifs sans référence.

    python pa_par_modele.py FOURNISSEUR_B
    python pa_par_modele.py FOURNISSEUR_E FOURNISSEUR_B
    python pa_par_modele.py --tous      tous ceux qui ont un tarif lu et
                                        des articles qui n'y sont pas

    Sans argument, c'est le FOURNISSEUR_B — et rien d'autre. Le script le
    dit maintenant en clair : ce défaut ressemble à un balayage complet sans
    en être un.

POURQUOI UN TROISIÈME OUTIL
    Chez le Fournisseur B, les deux côtés ne décrivent pas au même niveau.
    Le WMS liste des déclinaisons finies — « GIR1/GIR2 COQUILLE MODELEX4
    PRALINE T44CM P44CM » — quand le tarif liste des configurations —
    « MODELEX1-30° - Exécution standard - Dossier inclinable, accoudoirs
    X03 ».

    La similarité de libellé n'y peut rien : sur 174 articles elle a produit
    UN candidat, et il était faux (une option tablette rapprochée d'un
    fauteuil « sans tablette »). Ce n'est pas un défaut de seuil, c'est que
    les libellés ne sont pas comparables.

    Ce qui EST comparable, c'est le nom du modèle : MODELEX1, MODELEX2,
    MODELEX3, MODELEX4, MODELEX5, MODELEX6, MODELEX7. C'est la leçon du
    Fournisseur F, consignée dans CLAUDE.md §6 — « c'est le nom du modèle
    qui identifie ».

CE QUE ÇA PRODUIT
    Une vue d'arbitrage : pour chaque article du WMS, les lignes du tarif qui
    partagent son modèle, avec leurs prix et le rapport au DPA. Plusieurs
    candidats par article, c'est normal et voulu : c'est un humain qui
    tranche, pas le script. Rien n'est injecté.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

# Mots qui ne désignent aucun modèle : vocabulaire de catalogue, matières,
# couleurs, options. Sans ce filtre, « STANDARD » rapprocherait tout de tout.
BANALS = {
    "AVEC", "SANS", "POUR", "ET", "OU", "DE", "DU", "LA", "LE", "LES", "DES",
    "STANDARD", "EXECUTION", "VERSION", "COMPLET", "COMPLETE", "OPTION",
    "OPTIONS", "MONTEES", "MONTEE", "DOSSIER", "FIXE", "INCLINABLE", "ROUES",
    "ROUE", "AVANT", "ARRIERE", "ARRIERES", "POUCES", "POUCE", "KIT", "CONFIGURE",
    "CONFIGURES", "ACCOUDOIRS", "ACCOUDOIR", "PAIRE", "GAUCHE", "DROITE",
    "DROIT", "NOIR", "NOIRE", "BLANC", "BLEU", "ROUGE", "VERT", "GRIS", "GRISE",
    "BEIGE", "TAUPE", "MARRON", "ROSE", "JAUNE", "ORANGE", "VIOLET", "PRUNE",
    "PRALINE", "ARGENT", "CARBONE", "CUIR", "ARTIFICIEL", "BRUN", "FONCE",
    "CLAIR", "MAX", "MIN", "CM", "MM", "KG", "TAILLE", "LARGEUR", "PROFONDEUR",
    "SIEGE", "ASSISE", "NOUVEAU", "NOUVELLE", "JUSQU", "EPUISEMENT", "STOCK",
    "TARIF", "PRIX", "REMISE", "PUBLIC", "NET", "HT", "TTC", "LPP", "DMC",
    "GIR1", "GIR2", "GIR", "XXL", "XL", "XS", "SPORTY", "LIMITED", "EDITION",
    "ADULTE", "ADULTES", "ENFANT", "ENFANTS", "ADOLESCENT", "BARIATRIQUE",
    # Matières et finitions : elles se retrouvent sur des modèles sans
    # rapport. « SKAÏ » a fait matcher un REPOS MODELEX8 avec un MODELEX9,
    # au seul motif que les deux sont proposés en skaï.
    "SKAI", "TISSU", "BOIS", "MOUSSE", "POLYESTER", "COTON", "VELOURS",
    # Le public visé n'identifie aucun produit : sans ce filtre, une
    # « CHAUSSURE ORTHOPEDIQUE FEMME » se rapprochait d'une « SOCQUETTE
    # FEMME » par ce seul mot.
    "FEMME", "FEMMES", "HOMME", "HOMMES", "DAME", "DAMES", "MIXTE",
    "JUNIOR", "SENIOR", "UNISEXE",
    # Les suffixes de gamme (XTRA, NEW, EVO…) ne sont NI des modèles NI des
    # mots vides : ils qualifient un modèle et les distinguent entre eux.
    # Les bannir rapprochait « NEW MODELEX10 » de « MODELEX10 XTRA » ; les garder
    # comme modèles rapprochait « MODELEX11 XTRA » de « MODELEX12 XTRA ». Ils sont
    # donc traités comme des DISCRIMINANTS — voir SUFFIXES_GAMME.
    "ELECTRIQUE", "MANUEL", "MANUELS", "MANUELLE", "FAUTEUIL", "FAUTEUILS",
    "CHAISE", "LIT", "LITS", "ROLLATOR", "CANNE", "BEQUILLE", "SCOOTER",
    "TRICYCLE", "COQUILLE", "REPOS", "TABLETTE", "SEAU", "COUVERCLE",
    "TOILETTE", "BAIN", "TRANSFERT", "REGLABLE", "PLIANT", "PROTEGE",
    "VETEMENTS", "RELEVABLES", "RELEVABLE", "SURBAISSE", "INITIAL", "LIGHT",
    "EVO", "HEM2", "NIVEAUX", "METAL", "BARRIERES", "BARRIERE",
    # « DOUBLE » qualifie n'importe quoi — un stéthoscope À DOUBLE pavillon,
    # un tissu DOUBLÉ (doublure) — sans jamais identifier UN produit. C'est
    # devenu le seul mot commun entre un « STETHOSCOPE DOUBLE PAVILLON
    # MODELEX14 » et un « Protège-jambes doublé » chez le Fournisseur A,
    # decouvert le 17/09 en relisant les candidats « à relire » — un
    # stethoscope aurait recu le prix d'un protege-jambes.
    "DOUBLE", "DOUBLES", "SIMPLE", "SIMPLES", "UNIQUE", "UNIQUES",
    # Contenants et accessoires : ils décrivent l'emballage, jamais le
    # produit. Sans ce filtre, « MODELEX15 2% BIDON DE 5L » se
    # rapprochait d'un « ROBINET DE VIDANGE POUR BIDON 5L » par le seul
    # mot BIDON, et le classait DEVANT le vrai MODELEX15. Même chose pour
    # « MODELEX16 500 ML + POMPE », rapproché d'un support mural.
    "POMPE", "POMPES", "PPE", "BIDON", "BIDONS", "FLACON", "FLACONS",
    "SACHET", "SACHETS", "SUPPORT", "SUPPORTS", "DISTRIBUTEUR",
    "DISTRIBUTEURS", "DOSEUR", "DOSEURS", "ROBINET", "PICHET",
    "BOUTEILLE", "BOUTEILLES", "POCHE", "POCHES", "TUBE", "TUBES",
    "RECHARGE", "RECHARGES", "CARTOUCHE", "CARTOUCHES", "LINGETTE",
    "LINGETTES", "GACHETTE", "PULVERISATEUR", "VIDANGE",
}

# Un jeton qui n'est QU'une contenance — « 500ML », « 300ML », « 5L ».
#
# La contenance est déjà un discriminant : `contenances_compatibles` sait
# que « 12X0,5L » et « 500 ML » désignent la même unité. Mais elle servait
# AUSSI de clé de rapprochement, et c'est exactement l'erreur que ce module
# combat ailleurs pour les suffixes de gamme : un trait qui SÉPARE ne doit
# jamais RAPPROCHER. « MODELEX16 500 ML » tombait ainsi sur un
# « DISTRIBUTEUR 500 ML » et sur un « SUPPORT SOLUTION HYDROALCO 500ML »,
# tandis que le vrai MODELEX16 — que le tarif écrit « 12X0,5L », donc sans
# le jeton « 500ML » — sortait du classement.
EST_CONTENANCE = re.compile(
    r"^\d{1,5}(?:[.,]\d{1,3})?(?:ML|CL|LITRES?|L|KG|GR?|MM|CM)$")

# Un modèle : au moins trois caractères, en capitales, éventuellement avec
# des chiffres. « MODELEX1 », « MODELEX2 », « MODELEX3 », « 2215 ».
MOT = re.compile(r"[A-Z0-9][A-Z0-9°'-]{2,}")


def modeles(libelle, marque_exclue: set | None = None) -> set:
    """Les noms de modèle que porte un libellé.

    `marque_exclue` retire les mots du nom du fournisseur. Sans cela, le nom
    du Fournisseur J rapproche tout de ce fournisseur de n'importe quoi
    d'autre du même fournisseur : un « MODELEX17 » recevait le prix des
    « BANDELETTES MODELEX18 ».
    """
    if libelle is None or (isinstance(libelle, float) and pd.isna(libelle)):
        return set()
    import unicodedata
    texte = unicodedata.normalize("NFKD", str(libelle).upper())
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    exclus = marque_exclue or set()
    trouves = set()
    for mot in MOT.findall(texte):
        net = mot.strip("-'°")
        # Un suffixe de gamme ne RAPPROCHE pas — il ne fait que séparer, via
        # declinaison_compatible. Le laisser ici appariait « MODELEX13 XTRA »
        # à « MODELEX12 XTRA », avec un rapport de 1,006 que rien ne dénonçait.
        if (len(net) < 3 or net in BANALS or net in exclus
                or net in SUFFIXES_GAMME or EST_CONTENANCE.match(net)):
            continue
        # Un nombre seul n'identifie un modèle que s'il est long (2215).
        if net.isdigit() and len(net) < 4:
            continue
        trouves.add(net)
        # Un modèle composé porte aussi son radical : le tarif écrit
        # « MODELEX1-30° » d'un seul tenant quand le WMS écrit « MODELEX1 30° » en
        # deux mots. Sans ce découpage, les deux ne se rencontrent jamais —
        # et l'article du WMS tombait alors sur la seule ligne « MODELEX1 », celle
        # du dossier fixe à 200 € au lieu de 250 € (valeurs d'exemple).
        radical = re.split(r"[-']", net)[0]
        if len(radical) >= 3 and radical not in BANALS and not (
                radical.isdigit() and len(radical) < 4):
            trouves.add(radical)

    # Le recollage produit lui aussi des contenances — « 500 ML » écrit en
    # deux mots redevient « 500ML ». Même règle : une contenance sépare,
    # elle ne rapproche pas. Sans ce filtre, « MODELEX16 500 ML » tombait
    # sur un « DISTRIBUTEUR 500 ML NM » par ce seul jeton.
    trouves |= {code for code in codes_recolles(texte, exclus)
                if not EST_CONTENANCE.match(code)}
    return trouves


# Longueur d'un code recollé : en deçà on recolle du vocabulaire, au-delà on
# recolle des phrases.
RECOLLE_MIN, RECOLLE_MAX = 4, 12


def codes_recolles(texte: str, exclus: set | None = None) -> set:
    """Les références que le tarif espace et que le WMS colle — ou l'inverse.

    Le Fournisseur G écrit « CHUT AB 2214 B » là où la fiche du WMS porte
    « AB2214B ».
    Aucune comparaison mot à mot ne les rapproche : il faut recoller les
    fragments. On ne recolle qu'une suite courte dont au moins un morceau est
    un nombre — « AB » + « 2214 » + « B » donne AB2214B, mais « CHAUSSURE
    ORTHOPEDIQUE FEMME » ne donne rien.
    """
    exclus = exclus or set()
    mots = [m for m in re.split(r"[^A-Z0-9]+", texte) if m]
    trouves = set()
    for debut in range(len(mots)):
        for longueur in (2, 3):
            fin = debut + longueur
            if fin > len(mots):
                continue
            bloc = mots[debut:fin]
            # Au moins un nombre, et pas plus d'un mot « long » : sinon on
            # colle deux mots ordinaires et on fabrique un faux code.
            if not any(m.isdigit() for m in bloc):
                continue
            if sum(1 for m in bloc if len(m) > 4 and not m.isdigit()) > 0:
                continue
            colle = "".join(bloc)
            if (RECOLLE_MIN <= len(colle) <= RECOLLE_MAX
                    and not colle.isdigit()
                    and colle not in BANALS and colle not in exclus):
                trouves.add(colle)
    return trouves


# --------------------------------------------------------------------------
# Traits qui distinguent deux déclinaisons du MÊME modèle
#
# Le nom du modèle rapproche ; ces traits-là séparent. Un « MODELEX1 30° DI »
# (dossier inclinable, 250 €) et un « MODELEX1 DF » (dossier fixe, 200 €) —
# valeurs d'exemple — portent
# le même modèle et ne coûtent pas le même prix : sans ces traits, le premier
# recevait le prix du second, soit 20 % de moins.
#
# Chaque trait est une famille de synonymes. Deux valeurs différentes de la
# même famille, présentes des deux côtés, REJETTENT le candidat.

TRAITS = {
    "dossier": {
        "INCLINABLE": (r"\bDI\b", r"INCLINABLE", r"\b30\s*°", r"-30°", r"\b30 ?DEGRE"),
        "FIXE": (r"\bDF\b", r"DOSSIER FIXE", r"\bFIXE\b"),
    },
    "gamme": {
        "INITIAL": (r"\bINITIAL\b",),
        "STANDARD": (r"\bSTANDARD\b", r"EXECUTION STANDARD"),
    },
    "nomenclature": {
        "FMP": (r"\bFMP\b", r"NON-MODULAIRE", r"NON MODULAIRE"),
        "FRM": (r"\bFRM\b", r"\bMODULAIRE\b"),
    },
    "tablette": {
        "AVEC": (r"AVEC TABLETTE", r"COMPLET AVEC TABLETTE"),
        "SANS": (r"SANS TABLETTE",),
    },
    # Une OPTION n'est pas le produit qu'elle équipe. « OPTION TABLETTE
    # COQUILLE MODELEX4 » recevait le prix de la coquille complète —
    # 690,00 € contre un DPA de 15 € (valeurs d'exemple), soit quarante-six
    # fois trop. De même pour « OPT MODELEX1 REPOSE JAMBE », qui héritait des
    # 200 € du fauteuil (valeur d'exemple).
    "nature": {
        "OPTION": (r"\bOPTION\b", r"\bOPTIONS\b", r"\bOPT\b",
                   r"\bACCESSOIRE", r"\bPIECE DETACHEE", r"\bRECHANGE\b"),
        "PRODUIT": (r"EXECUTION STANDARD", r"COMPLET\b", r"DOSSIER FIXE",
                    r"DOSSIER INCLINABLE"),
    },
}


# Contenance unitaire, ramenée au millilitre (ou au gramme). Chez les
# fournisseurs de détergent, c'est LE discriminant : « MODELEX17 85 100ML » et
# « MODELEX17 85 6X300ML » portent le même modèle et ne sont pas le même
# produit. Le tarif écrit « 12X0,5L » là où le WMS écrit « 500 ML » — il faut
# donc diviser par le conditionnement et convertir avant de comparer.
_UNITES = {"ML": 1.0, "CL": 10.0, "L": 1000.0, "LITRE": 1000.0,
           "LITRES": 1000.0, "G": 1.0, "GR": 1.0, "KG": 1000.0}

_CONTENANCE = re.compile(
    r"(?:(\d{1,4})\s*[xX*]\s*)?"              # conditionnement facultatif
    r"(\d{1,5}(?:[.,]\d{1,3})?)\s*"           # la valeur
    r"(ML|CL|LITRES?|L|KG|GR?)\b"             # l'unité
)


def contenance(libelle) -> float | None:
    """La contenance d'UNE unité, en millilitres — ou None.

    « 12X0,5L » vaut 500 : on ignore le 12, qui est le conditionnement, et on
    convertit le 0,5 L. « 500 ML » vaut 500 aussi. Les deux se comparent alors.
    """
    if libelle is None or (isinstance(libelle, float) and pd.isna(libelle)):
        return None
    import unicodedata
    texte = unicodedata.normalize("NFKD", str(libelle).upper())
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    for trouve in _CONTENANCE.finditer(texte):
        _, valeur, unite = trouve.groups()
        facteur = _UNITES.get(unite)
        if not facteur:
            continue
        try:
            nombre = float(valeur.replace(",", "."))
        except ValueError:
            continue
        millilitres = nombre * facteur
        # Une contenance plausible pour un produit d'hygiène : entre 10 ml et
        # 30 litres. Au-delà, c'est une dimension ou une référence.
        if 10 <= millilitres <= 30000:
            return round(millilitres, 2)
    return None


def contenances_compatibles(a, b) -> bool:
    """Deux contenances connues doivent être égales, à 2 % près."""
    if a is None or b is None:
        return True
    return abs(a - b) <= 0.02 * max(a, b)


def traits(libelle) -> dict:
    """Les traits distinctifs que porte un libellé."""
    import unicodedata
    if libelle is None or (isinstance(libelle, float) and pd.isna(libelle)):
        return {}
    texte = unicodedata.normalize("NFKD", str(libelle).upper())
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    trouves = {}
    for famille, valeurs in TRAITS.items():
        for valeur, motifs in valeurs.items():
            if any(re.search(m, texte) for m in motifs):
                # « FIXE » ne doit pas l'emporter sur « INCLINABLE » quand le
                # libellé porte les deux (« dossier fixe ou inclinable ») :
                # le premier trouvé gagne, et l'ordre du dict le fixe.
                trouves.setdefault(famille, valeur)
    return trouves


def traits_compatibles(a: dict, b: dict) -> bool:
    """Deux libellés peuvent-ils désigner le même produit ?

    Un trait absent d'un côté ne rejette pas — les deux catalogues ne disent
    pas tout. Un trait présent des DEUX côtés et différent, si.
    """
    for famille in set(a) & set(b):
        if a[famille] != b[famille]:
            return False
    return True


# La TAILLE, isolée du reste des nombres. « MODELEX19 A3 … T3 » porte deux
# nombres : le 3 de la gamme A3 et le 3 de la taille T3. Les comparer en vrac
# ne sépare rien — « A3 … T3 » et « A3 … T1 » partagent toujours le 3 de la
# gamme, et cinq genouillères de tailles différentes recevaient le prix du T1.
TAILLE = re.compile(r"\bT\.?\s?(\d{1,2})\b|\bTAILLE\s*:?\s*(\d{1,2})\b"
                    r"|\bG\.\s?(\d{1,2})\b")

# Le côté, y compris dans ses formes abrégées. Le Fournisseur E écrit « D. » là où
# le WMS écrit « DROITE » : sans cette équivalence, une genouillère GAUCHE
# recevait le prix d'une DROITE sans que rien ne s'y oppose.
COTE_DROIT = re.compile(r"\bDROITE?\b|\bD\.\s|\bDROIT\b")
COTE_GAUCHE = re.compile(r"\bGAUCHES?\b|\bG\.\s")


def _propre(libelle) -> str:
    import unicodedata
    t = unicodedata.normalize("NFKD", str(libelle or "").upper())
    return "".join(c for c in t if not unicodedata.combining(c))


def taille(libelle) -> set:
    """Les tailles déclarées par un libellé, en nombres."""
    trouves = set()
    for groupes in TAILLE.findall(_propre(libelle)):
        for valeur in groupes:
            if valeur:
                trouves.add(int(valeur))
    return trouves


# Les tailles qui s'écrivent en LETTRES. « MODELEX20 M » et « MODELEX20 L » ne sont
# pas le même rollator, mais ils partagent leur coloris — et une simple
# intersection de discriminants les laissait passer, parce que « ROUGE » était
# commun aux deux. Une taille se compare en égalité, jamais en recouvrement.
TAILLES_LETTRES = {"XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL",
                   "TU", "TS", "TM", "TL", "TXL"}


# Suffixes qui déclinent une gamme. Ils doivent être IDENTIQUES des deux
# côtés : « NEW MODELEX10 » et « MODELEX10 XTRA » partagent leur modèle et ne sont
# pas la même chaussure, pas plus que « MODELEX11 XTRA » et « MODELEX12 XTRA ».
SUFFIXES_GAMME = {"XTRA", "EXTRA", "NEW", "EVO", "MAXI", "MINI", "ULTRA",
                  "PREMIUM", "CLASSIC", "PLUS", "LIGHT", "SOFT", "PRO",
                  "INITIAL", "STANDARD"}


def suffixes(libelle) -> set:
    """Les suffixes de gamme portés par un libellé."""
    mots = re.split(r"[^A-Z0-9]+", _propre(libelle))
    return {m for m in mots if m in SUFFIXES_GAMME}


def taille_lettre(libelle) -> set:
    """Les tailles en lettres, isolées comme mots entiers."""
    mots = re.split(r"[^A-Z0-9]+", _propre(libelle))
    return {m for m in mots if m in TAILLES_LETTRES}


def cote(libelle) -> str | None:
    """« DROIT », « GAUCHE », ou None si le libellé ne le dit pas."""
    texte = _propre(libelle) + " "
    droit = bool(COTE_DROIT.search(texte))
    gauche = bool(COTE_GAUCHE.search(texte))
    if droit and not gauche:
        return "DROIT"
    if gauche and not droit:
        return "GAUCHE"
    return None


def declinaison_compatible(libelle_wms, libelle_tarif, marque=None) -> bool:
    """La TAILLE et le CÔTÉ doivent être identiques, pas ressemblants.

    Le nom de modèle rapproche une gamme entière ; il ne distingue pas ses
    déclinaisons. Sans ce contrôle, cinq « GENOUILLERE MODELEX19 A3 » de
    tailles 2 à 6 recevaient toutes le prix du T1, et une GAUCHE recevait
    celui d'une DROITE — c'est l'erreur que le brief décrit mot pour mot.

    On réutilise par_libelle, qui porte déjà ce savoir et l'a payé cher :
    `nombres()` pour les tailles et calibres, `discriminants()` pour le côté,
    les coloris et les mots de gamme.
    """
    import par_libelle

    # La taille d'abord, isolée des autres nombres : c'est elle qui sépare les
    # déclinaisons d'une même gamme, et elle seule.
    t_wms, t_tarif = taille(libelle_wms), taille(libelle_tarif)
    if t_wms and t_tarif and not (t_wms & t_tarif):
        return False

    # Les suffixes de gamme, en ÉGALITÉ stricte.
    s_wms, s_tarif = suffixes(libelle_wms), suffixes(libelle_tarif)
    if s_wms != s_tarif:
        return False

    # Les tailles en lettres, en ÉGALITÉ stricte : S, M, L, XL séparent des
    # produits que tout le reste du libellé confond.
    l_wms, l_tarif = taille_lettre(libelle_wms), taille_lettre(libelle_tarif)
    if l_wms and l_tarif and l_wms != l_tarif:
        return False

    # Le côté ensuite : une gauche n'est pas une droite, à aucun prix.
    c_wms, c_tarif = cote(libelle_wms), cote(libelle_tarif)
    if c_wms and c_tarif and c_wms != c_tarif:
        return False

    # La CONTENANCE, comparée en millilitres : c'est le discriminant qui
    # sépare vraiment deux conditionnements d'un même produit.
    v_wms, v_tarif = contenance(libelle_wms), contenance(libelle_tarif)
    if not contenances_compatibles(v_wms, v_tarif):
        return False

    # Les nombres bruts, en dernier recours SEULEMENT.
    #
    # Ils ne savent pas convertir : « MODELEX16 500 ML » porte {500} quand le
    # tarif « MODELEX16 NPC HF 12X0,5L » porte {12 ; 0,5}. Disjoints, donc
    # rejet — alors que les deux désignent la MÊME unité de 500 ml, ce que la
    # contenance vient d'établir trois lignes plus haut. Le contrôle grossier
    # contredisait le contrôle fin, et c'est le grossier qui l'emportait :
    # MODELEX16 restait introuvable alors que son prix était au tarif.
    #
    # Quand la contenance a tranché des deux côtés, elle fait foi et on ne
    # redemande rien aux nombres bruts.
    if v_wms is None or v_tarif is None:
        n_wms = par_libelle.nombres(libelle_wms)
        n_tarif = par_libelle.nombres(libelle_tarif)
        if n_wms and n_tarif and not (n_wms & n_tarif):
            return False

    d_wms = par_libelle.discriminants(libelle_wms, marque)
    d_tarif = par_libelle.discriminants(libelle_tarif)
    # Présent d'un seul côté : on ne rejette pas — les deux catalogues
    # n'écrivent pas la même chose. Présents des deux côtés et disjoints :
    # ce sont deux produits.
    if d_wms and d_tarif and not (d_wms & d_tarif):
        return False
    return True


def proposer(motif: str) -> pd.DataFrame:
    import pa_completer as P
    import pa_conditionnement as C

    articles = P.articles_du_perimetre(motif)
    if articles.empty:
        return pd.DataFrame()
    noms = tuple(sorted(articles["Nom fournisseur"].dropna().unique()))
    tarif = P.tarif_du_fournisseur(motif, noms)
    if tarif.empty or "designation" not in tarif.columns:
        return pd.DataFrame()
    tarif = tarif.dropna(subset=["designation"]).copy()

    # La marque du fournisseur ne distingue rien : elle est des deux côtés, sur
    # toutes les lignes. On la retire, ainsi que tout mot présent sur plus de la
    # moitié des lignes du tarif — chez le Fournisseur J, « POMPE » revient partout et
    # rapprochait un MODELEX16 d'une « POMPE 25CC CAPS ROUGE ».
    marque = set()
    for nom in noms:
        marque |= {m for m in re.split(r"[^A-Z0-9]+", str(nom).upper())
                   if len(m) >= 3}
    frequents = {}
    for designation in tarif["designation"]:
        for mot in set(modeles(designation)):
            frequents[mot] = frequents.get(mot, 0) + 1
    seuil = max(3, len(tarif) // 2)
    marque |= {mot for mot, n in frequents.items() if n >= seuil}

    tarif["modeles"] = tarif["designation"].map(lambda x: modeles(x, marque))
    tarif["traits"] = tarif["designation"].map(traits)
    tarif["contenance"] = tarif["designation"].map(contenance)

    pa = P.pa_actuels().set_index("code_article")["pa_wms"]

    lignes = []
    for _, art in articles.iterrows():
        siens = modeles(art["Libellé déclinaison ^(1)"], marque)
        if not siens:
            continue
        code = str(art["Code article"]).strip()
        dpa = pa.get(code)
        mes_traits = traits(art["Libellé déclinaison ^(1)"])
        ma_contenance = contenance(art["Libellé déclinaison ^(1)"])
        candidats = []
        for _, ligne in tarif.iterrows():
            communs = siens & ligne["modeles"]
            if not communs:
                continue
            # Même modèle ne veut pas dire même produit : un dossier
            # inclinable n'est pas un dossier fixe, et ils ne coûtent pas
            # le même prix.
            if not traits_compatibles(mes_traits, ligne["traits"]):
                continue
            # La contenance sépare deux flacons du même produit.
            if not contenances_compatibles(ma_contenance, ligne["contenance"]):
                continue
            # La taille et le côté séparent deux déclinaisons d'une gamme.
            if not declinaison_compatible(art["Libellé déclinaison ^(1)"],
                                          ligne["designation"],
                                          art["Nom fournisseur"]):
                continue
            candidats.append((communs, ligne))
        if not candidats:
            continue
        # Le meilleur d'abord : le plus de modèles en commun, puis le prix le
        # plus proche du DPA quand il existe.
        def tri(c):
            communs, ligne = c
            prix = ligne.get("prix_achat_unitaire_ht")
            ecart = 99.0
            if dpa and prix and dpa > 0:
                r = float(prix) / float(dpa)
                ecart = abs(r - 1)
            # Règle d'achat donnée par les appros : sur les coquilles, on
            # prend le modèle COMPLET AVEC TABLETTE. Quand le WMS ne précise
            # rien, la version avec tablette passe donc devant.
            tablette = ligne["traits"].get("tablette")
            sans_tablette = 1 if (tablette == "SANS"
                                  and "tablette" not in mes_traits) else 0
            return (-len(communs), sans_tablette, ecart)

        candidats.sort(key=tri)
        for rang, (communs, ligne) in enumerate(candidats[:4], 1):
            prix = ligne.get("prix_achat_unitaire_ht")
            # Le prix du tarif peut porter sur un conditionnement — chez
            # le Fournisseur J, « 12X0,5L » vaut 60,00 € pour douze flacons
            # (valeur d'exemple). Sans la division, le rapport au DPA vaut
            # douze fois trop et le candidat
            # paraît absurde alors qu'il est bon.
            diviseur, source_div = 1.0, ""
            lu, nom = C.diviseur_depuis_designation(
                ligne["designation"], art["Libellé déclinaison ^(1)"])
            if lu:
                diviseur, source_div = float(lu), nom
            unitaire = (round(float(prix) / diviseur, 4)
                        if prix and diviseur else prix)
            rapport = None
            if dpa and unitaire and float(dpa) > 0:
                rapport = round(float(unitaire) / float(dpa), 3)
            lignes.append({
                "Code article": code,
                "Libellé WMS": art["Libellé déclinaison ^(1)"],
                "Rang": rang,
                "Modèles communs": " + ".join(sorted(communs)),
                "Libellé tarif": ligne["designation"],
                "Prix tarif": prix,
                "Diviseur": diviseur if diviseur > 1 else None,
                "Source diviseur": source_div or None,
                "Prix unitaire": unitaire,
                "Paliers": ligne.get("paliers"),
                "Ancien PA": dpa,
                "Rapport": rapport,
                "Candidats pour cet article": len(candidats),
                "Fichier source": ligne.get("fichier_tarif"),
            })
    res = pd.DataFrame(lignes)
    if res.empty:
        return res

    # Un facteur SYSTÉMATIQUE sur tout le fournisseur — le tarif du Fournisseur G cote
    # la paire quand le WMS compte les souliers un par un, et les rapports se
    # groupent alors sur 2,00–2,05. Sans cette passe, des candidats justes
    # sortent de la fourchette de crédibilité et passent pour douteux.
    #
    # On réutilise la déduction de pa_conditionnement, avec sa garantie : le
    # même entier sur au moins cinq articles, et sur plus de 60 % d'entre eux.
    rang1 = res[res["Rang"] == 1]
    rapports = pd.to_numeric(rang1["Rapport"], errors="coerce").dropna()
    deduit, preuve, _ = C.deduire_par_systematicite(rapports)
    if deduit:
        colle = res["Rapport"].map(lambda r: C.suit_le_facteur(r, deduit))
        sans_dpa = res["Rapport"].isna()
        vise = colle | sans_dpa
        res.loc[vise, "Prix unitaire"] = (
            pd.to_numeric(res.loc[vise, "Prix unitaire"], errors="coerce")
            / float(deduit)).round(4)
        res.loc[vise, "Rapport"] = (
            pd.to_numeric(res.loc[vise, "Prix unitaire"], errors="coerce")
            / pd.to_numeric(res.loc[vise, "Ancien PA"], errors="coerce")).round(3)
        res.loc[vise, "Diviseur"] = float(deduit)
        res.loc[vise, "Source diviseur"] = "déduit du fournisseur"
        print(f"    facteur /{deduit} appliqué à {int(vise.sum())} lignes — {preuve}")
    return res


def fournisseurs_a_traiter() -> list[str]:
    """Les fournisseurs qui ont un tarif ET des articles sans prix.

    C'est exactement la population que ce module vise : un tarif est là,
    il est lu, et pourtant des articles n'y sont pas retrouvés. Le motif
    « aucun tarif reçu » est exclu — sans tarif, il n'y a rien à
    rapprocher, et c'est un mail qu'il faut, pas un algorithme.
    """
    import warnings as _w

    _w.filterwarnings("ignore")
    from extracteurs.base import chemin_lisible
    from main import SORTIE_ACHATS

    etat = SORTIE_ACHATS / "pa_etat_par_article.xlsx"
    if not etat.exists():
        raise SystemExit("Lancez d'abord : python run_pa.py")
    d = pd.read_excel(chemin_lisible(etat), skiprows=3, dtype=str)
    vises = d[d["Motif"].fillna("").str.contains(
        "pas dans son tarif", case=False)]
    comptes = vises["Fournisseur"].dropna().value_counts()
    return [str(n) for n in comptes.index]


def main() -> None:
    import mise_en_forme
    from main import SORTIE_ACHATS

    # Le defaut historique est FOURNISSEUR_B, et il ne se devine pas : lancer le
    # script sans argument RESSEMBLE a un balayage complet, et n'en est pas
    # un. Je m'y suis laisse prendre le 16/09 en annoncant « mesure sur tout
    # le perimetre » ce qui n'etait qu'un passage sur le Fournisseur B. On le dit donc
    # a voix haute, et on offre le vrai balayage.
    if "--tous" in sys.argv:
        cibles = fournisseurs_a_traiter()
        print(f"--tous : {len(cibles)} fournisseurs ayant un tarif lu et des "
              f"articles qui n'y sont pas retrouvés")
    else:
        cibles = [a for a in sys.argv[1:] if not a.startswith("--")]
        if not cibles:
            cibles = ["FOURNISSEUR_B"]
            print("Aucun fournisseur donné — FOURNISSEUR_B par défaut.")
            print("Pour balayer tous les fournisseurs concernés : "
                  "python pa_par_modele.py --tous")
    tables = []
    for motif in cibles:
        print(f"--- {motif}")
        try:
            t = proposer(motif)
        except Exception as err:
            print(f"    ! {type(err).__name__} : {err}")
            continue
        if not t.empty:
            print(f"    {t['Code article'].nunique()} articles avec au moins "
                  f"un candidat, {len(t)} lignes de proposition")
        tables.append(t)

    utiles = [t for t in tables if not t.empty]
    if not utiles:
        print("aucun candidat")
        return
    tout = pd.concat(utiles, ignore_index=True)
    # Le rang 1 dont le rapport est crédible remonte : c'est là que la
    # relecture paie le plus vite.
    tout["_sur"] = (tout["Rang"] == 1) & tout["Rapport"].between(0.5, 2.0)
    tout = tout.sort_values(["_sur", "Code article", "Rang"],
                            ascending=[False, True, True]).drop(columns="_sur")

    chemin = mise_en_forme.chemin_ecriture(
        SORTIE_ACHATS / "pa_par_modele.xlsx")
    with pd.ExcelWriter(chemin, engine="openpyxl") as w:
        tout.to_excel(w, sheet_name="Par modèle", index=False, startrow=3)
    mise_en_forme.formater(chemin, {
        "Par modèle": {
            "ligne_entete": 4, "figer_colonne": 2,
            "titre": "Candidats par NOM DE MODÈLE — tarifs sans référence",
            "sous_titre": "Plusieurs lignes par article, c'est voulu : "
                          "l'arbitrage est humain. Rien n'est injecté."},
    })
    print(f"\n  -> {chemin}")


if __name__ == "__main__":
    main()
