# -*- coding: utf-8 -*-
"""
PA workstream: match by MODEL NAME, for price lists without references.

    python pa_par_modele.py FOURNISSEUR_B
    python pa_par_modele.py FOURNISSEUR_E FOURNISSEUR_B
    python pa_par_modele.py --tous      every supplier that has a price list
                                        read and articles missing from it

    With no argument it is FOURNISSEUR_B, and nothing else. The script now
    says so explicitly: that default looks like a full sweep without being
    one.

WHY A THIRD TOOL
    At Supplier B, the two sides do not describe things at the same level.
    The WMS lists finished variants - "GIR1/GIR2 COQUILLE MODELEX4 PRALINE
    T44CM P44CM" - while the price list lists configurations -
    "MODELEX1-30° - Exécution standard - Dossier inclinable, accoudoirs
    X03".

    Label similarity can do nothing about that: over 174 articles it
    produced ONE candidate, and it was wrong (a tray option matched to a
    chair described as "without tray"). This is not a threshold problem,
    it is that the labels are not comparable.

    What IS comparable is the model name: MODELEX1, MODELEX2, MODELEX3,
    MODELEX4, MODELEX5, MODELEX6, MODELEX7. That is the lesson of
    Supplier F, recorded in the project notes: "it is the model name
    that identifies".

WHAT IT PRODUCES
    An arbitration view: for each WMS article, the price list rows sharing
    its model, with their prices and their ratio to the DPA. Several
    candidates per article is normal and intended: a human decides, not the
    script. Nothing is injected.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

import pandas as pd

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

# Words that designate no model: catalogue vocabulary, materials, colours,
# options. Without this filter, "STANDARD" would match everything to
# everything.
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
    # Materials and finishes: they recur on models with nothing in common.
    # "SKAÏ" made a REPOS MODELEX8 match a MODELEX9, for the sole reason
    # that both are offered in leatherette.
    "SKAI", "TISSU", "BOIS", "MOUSSE", "POLYESTER", "COTON", "VELOURS",
    # The intended audience identifies no product: without this filter, a
    # "CHAUSSURE ORTHOPEDIQUE FEMME" matched a "SOCQUETTE FEMME" on that
    # single word.
    "FEMME", "FEMMES", "HOMME", "HOMMES", "DAME", "DAMES", "MIXTE",
    "JUNIOR", "SENIOR", "UNISEXE",
    # Range suffixes (XTRA, NEW, EVO...) are NEITHER models NOR stop words:
    # they qualify a model and tell models apart. Banning them matched
    # "NEW MODELEX10" to "MODELEX10 XTRA"; keeping them as models matched
    # "MODELEX11 XTRA" to "MODELEX12 XTRA". So they are treated as
    # DISCRIMINANTS - see SUFFIXES_GAMME.
    "ELECTRIQUE", "MANUEL", "MANUELS", "MANUELLE", "FAUTEUIL", "FAUTEUILS",
    "CHAISE", "LIT", "LITS", "ROLLATOR", "CANNE", "BEQUILLE", "SCOOTER",
    "TRICYCLE", "COQUILLE", "REPOS", "TABLETTE", "SEAU", "COUVERCLE",
    "TOILETTE", "BAIN", "TRANSFERT", "REGLABLE", "PLIANT", "PROTEGE",
    "VETEMENTS", "RELEVABLES", "RELEVABLE", "SURBAISSE", "INITIAL", "LIGHT",
    "EVO", "HEM2", "NIVEAUX", "METAL", "BARRIERES", "BARRIERE",
    # "DOUBLE" qualifies anything - a stethoscope with a DOUBLE head, a
    # LINED fabric ("doublé") - without ever identifying ONE product. It
    # became the only word shared by a "STETHOSCOPE DOUBLE PAVILLON
    # MODELEX14" and a "Protège-jambes doublé" at Supplier A, found on
    # 17/09 while reviewing the "to review" candidates - a stethoscope
    # would have been given the price of a leg protector.
    "DOUBLE", "DOUBLES", "SIMPLE", "SIMPLES", "UNIQUE", "UNIQUES",
    # Containers and accessories: they describe the packaging, never the
    # product. Without this filter, "MODELEX15 2% BIDON DE 5L" matched a
    # "ROBINET DE VIDANGE POUR BIDON 5L" through the single word BIDON,
    # and ranked it AHEAD of the real MODELEX15. Same for "MODELEX16
    # 500 ML + POMPE", matched to a wall bracket.
    "POMPE", "POMPES", "PPE", "BIDON", "BIDONS", "FLACON", "FLACONS",
    "SACHET", "SACHETS", "SUPPORT", "SUPPORTS", "DISTRIBUTEUR",
    "DISTRIBUTEURS", "DOSEUR", "DOSEURS", "ROBINET", "PICHET",
    "BOUTEILLE", "BOUTEILLES", "POCHE", "POCHES", "TUBE", "TUBES",
    "RECHARGE", "RECHARGES", "CARTOUCHE", "CARTOUCHES", "LINGETTE",
    "LINGETTES", "GACHETTE", "PULVERISATEUR", "VIDANGE",
}

# A token that is ONLY a capacity: "500ML", "300ML", "5L".
#
# Capacity is already a discriminant: `contenances_compatibles` knows that
# "12X0,5L" and "500 ML" designate the same unit. But it was ALSO serving
# as a matching key, and that is exactly the mistake this module fights
# elsewhere for range suffixes: a trait that SEPARATES must never BRING
# TOGETHER. "MODELEX16 500 ML" was thus landing on a "DISTRIBUTEUR 500 ML"
# and on a "SUPPORT SOLUTION HYDROALCO 500ML", while the real MODELEX16 -
# which the price list writes as "12X0,5L", hence without the "500ML"
# token - dropped out of the ranking.
EST_CONTENANCE = re.compile(
    r"^\d{1,5}(?:[.,]\d{1,3})?(?:ML|CL|LITRES?|L|KG|GR?|MM|CM)$")

# A model: at least three characters, in capitals, possibly with digits.
# "MODELEX1", "MODELEX2", "MODELEX3", "2215".
MOT = re.compile(r"[A-Z0-9][A-Z0-9°'-]{2,}")


def modeles(libelle, marque_exclue: set | None = None) -> set:
    """The model names a label carries.

    `marque_exclue` removes the words of the supplier's name. Without it,
    Supplier J's name matches anything from that supplier to anything else
    from the same supplier: a "MODELEX17" was given the price of
    "BANDELETTES MODELEX18".
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
        # A range suffix does not BRING TOGETHER - it only separates, via
        # declinaison_compatible. Leaving it here paired "MODELEX13 XTRA"
        # with "MODELEX12 XTRA", at a ratio of 1.006 that nothing flagged.
        if (len(net) < 3 or net in BANALS or net in exclus
                or net in SUFFIXES_GAMME or EST_CONTENANCE.match(net)):
            continue
        # A number on its own only identifies a model when it is long (2215).
        if net.isdigit() and len(net) < 4:
            continue
        trouves.add(net)
        # A compound model also carries its radical: the price list writes
        # "MODELEX1-30°" as one piece where the WMS writes "MODELEX1 30°"
        # in two words. Without this split the two never meet - and the WMS
        # article then landed on the only "MODELEX1" line, the one for the
        # fixed backrest at EUR 200 instead of EUR 250 (example values).
        radical = re.split(r"[-']", net)[0]
        if len(radical) >= 3 and radical not in BANALS and not (
                radical.isdigit() and len(radical) < 4):
            trouves.add(radical)

    # Re-gluing also produces capacities: "500 ML" written as two words
    # becomes "500ML" again. Same rule: a capacity separates, it does not
    # bring together. Without this filter, "MODELEX16 500 ML" landed on a
    # "DISTRIBUTEUR 500 ML NM" through that single token.
    trouves |= {code for code in codes_recolles(texte, exclus)
                if not EST_CONTENANCE.match(code)}
    return trouves


# Length of a re-glued code: below it we glue vocabulary together, above it
# we glue sentences.
RECOLLE_MIN, RECOLLE_MAX = 4, 12


def codes_recolles(texte: str, exclus: set | None = None) -> set:
    """The references the price list spaces out and the WMS glues, or vice
    versa.

    Supplier G writes "CHUT AB 2214 B" where the WMS record carries
    "AB2214B".
    No word-by-word comparison brings them together: the fragments have to
    be glued back. We only glue a short run in which at least one piece is
    a number - "AB" + "2214" + "B" gives AB2214B, but "CHAUSSURE
    ORTHOPEDIQUE FEMME" gives nothing.
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
            # At least one number, and no more than one "long" word:
            # otherwise we glue two ordinary words and manufacture a fake
            # code.
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
# Traits that tell apart two variants of the SAME model
#
# The model name brings together; these traits separate. A "MODELEX1 30° DI"
# (reclining backrest, EUR 250) and a "MODELEX1 DF" (fixed backrest,
# EUR 200) - example values - carry
# the same model and do not cost the same: without these traits, the first
# was given the price of the second, i.e. 20% less.
#
# Each trait is a family of synonyms. Two different values of the same
# family, present on both sides, REJECT the candidate.

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
    # An OPTION is not the product it equips. "OPTION TABLETTE COQUILLE
    # MODELEX4" was given the price of the complete shell - EUR 690.00
    # against a DPA of EUR 15 (example values), i.e. forty-six times too
    # much. Likewise for "OPT MODELEX1 REPOSE JAMBE", which inherited the
    # chair's EUR 200 (example value).
    "nature": {
        "OPTION": (r"\bOPTION\b", r"\bOPTIONS\b", r"\bOPT\b",
                   r"\bACCESSOIRE", r"\bPIECE DETACHEE", r"\bRECHANGE\b"),
        "PRODUIT": (r"EXECUTION STANDARD", r"COMPLET\b", r"DOSSIER FIXE",
                    r"DOSSIER INCLINABLE"),
    },
}


# Unit capacity, brought back to the millilitre (or the gram). At detergent
# suppliers this is THE discriminant: "MODELEX17 85 100ML" and "MODELEX17
# 85 6X300ML" carry the same model and are not the same product. The price
# list writes "12X0,5L" where the WMS writes "500 ML" - so the pack size
# must be divided out and the units converted before comparing.
_UNITES = {"ML": 1.0, "CL": 10.0, "L": 1000.0, "LITRE": 1000.0,
           "LITRES": 1000.0, "G": 1.0, "GR": 1.0, "KG": 1000.0}

_CONTENANCE = re.compile(
    r"(?:(\d{1,4})\s*[xX*]\s*)?"              # optional pack size
    r"(\d{1,5}(?:[.,]\d{1,3})?)\s*"           # the value
    r"(ML|CL|LITRES?|L|KG|GR?)\b"             # the unit
)


def contenance(libelle) -> float | None:
    """The capacity of ONE unit, in millilitres - or None.

    "12X0,5L" is 500: we ignore the 12, which is the pack size, and convert
    the 0.5 L. "500 ML" is 500 as well. The two can then be compared.
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
        # A plausible capacity for a hygiene product: between 10 ml and
        # 30 litres. Beyond that it is a dimension or a reference.
        if 10 <= millilitres <= 30000:
            return round(millilitres, 2)
    return None


def contenances_compatibles(a, b) -> bool:
    """Two known capacities must be equal, to within 2%."""
    if a is None or b is None:
        return True
    return abs(a - b) <= 0.02 * max(a, b)


def traits(libelle) -> dict:
    """The distinguishing traits a label carries."""
    import unicodedata
    if libelle is None or (isinstance(libelle, float) and pd.isna(libelle)):
        return {}
    texte = unicodedata.normalize("NFKD", str(libelle).upper())
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    trouves = {}
    for famille, valeurs in TRAITS.items():
        for valeur, motifs in valeurs.items():
            if any(re.search(m, texte) for m in motifs):
                # "FIXE" must not win over "INCLINABLE" when the label
                # carries both ("dossier fixe ou inclinable"): the first
                # found wins, and the dict order fixes that.
                trouves.setdefault(famille, valeur)
    return trouves


def traits_compatibles(a: dict, b: dict) -> bool:
    """Can two labels designate the same product?

    A trait absent from one side does not reject - the two catalogues do
    not state everything. A trait present on BOTH sides and different does.
    """
    for famille in set(a) & set(b):
        if a[famille] != b[famille]:
            return False
    return True


# The SIZE, isolated from the other numbers. "MODELEX19 A3 ... T3" carries
# two numbers: the 3 of the A3 range and the 3 of size T3. Comparing them in
# bulk separates nothing - "A3 ... T3" and "A3 ... T1" always share the 3 of
# the range, and five knee braces in different sizes were all given the
# price of the T1.
TAILLE = re.compile(r"\bT\.?\s?(\d{1,2})\b|\bTAILLE\s*:?\s*(\d{1,2})\b"
                    r"|\bG\.\s?(\d{1,2})\b")

# The side, including its abbreviated forms. Supplier E writes "D." where
# the WMS writes "DROITE": without this equivalence, a LEFT knee brace was
# given the price of a RIGHT one with nothing standing in the way.
COTE_DROIT = re.compile(r"\bDROITE?\b|\bD\.\s|\bDROIT\b")
COTE_GAUCHE = re.compile(r"\bGAUCHES?\b|\bG\.\s")


def _propre(libelle) -> str:
    import unicodedata
    t = unicodedata.normalize("NFKD", str(libelle or "").upper())
    return "".join(c for c in t if not unicodedata.combining(c))


def taille(libelle) -> set:
    """The sizes a label declares, as numbers."""
    trouves = set()
    for groupes in TAILLE.findall(_propre(libelle)):
        for valeur in groupes:
            if valeur:
                trouves.add(int(valeur))
    return trouves


# Sizes written as LETTERS. "MODELEX20 M" and "MODELEX20 L" are not the same
# rollator, but they share their colour - and a plain intersection of
# discriminants let them through, because "ROUGE" was common to both. A size
# is compared by equality, never by overlap.
TAILLES_LETTRES = {"XXS", "XS", "S", "M", "L", "XL", "XXL", "XXXL",
                   "TU", "TS", "TM", "TL", "TXL"}


# Suffixes that decline a range. They must be IDENTICAL on both sides:
# "NEW MODELEX10" and "MODELEX10 XTRA" share their model and are not the
# same shoe, any more than "MODELEX11 XTRA" and "MODELEX12 XTRA" are.
SUFFIXES_GAMME = {"XTRA", "EXTRA", "NEW", "EVO", "MAXI", "MINI", "ULTRA",
                  "PREMIUM", "CLASSIC", "PLUS", "LIGHT", "SOFT", "PRO",
                  "INITIAL", "STANDARD"}


def suffixes(libelle) -> set:
    """The range suffixes a label carries."""
    mots = re.split(r"[^A-Z0-9]+", _propre(libelle))
    return {m for m in mots if m in SUFFIXES_GAMME}


def taille_lettre(libelle) -> set:
    """Letter sizes, isolated as whole words."""
    mots = re.split(r"[^A-Z0-9]+", _propre(libelle))
    return {m for m in mots if m in TAILLES_LETTRES}


def cote(libelle) -> str | None:
    """Returns "DROIT", "GAUCHE", or None if the label does not say."""
    texte = _propre(libelle) + " "
    droit = bool(COTE_DROIT.search(texte))
    gauche = bool(COTE_GAUCHE.search(texte))
    if droit and not gauche:
        return "DROIT"
    if gauche and not droit:
        return "GAUCHE"
    return None


def declinaison_compatible(libelle_wms, libelle_tarif, marque=None) -> bool:
    """SIZE and SIDE must be identical, not similar.

    The model name brings a whole range together; it does not tell its
    variants apart. Without this check, five "GENOUILLERE MODELEX19 A3" in
    sizes 2 to 6 were all given the price of the T1, and a LEFT one was
    given that of a RIGHT one - the very error the brief describes word for
    word.

    We reuse par_libelle, which already holds this knowledge and paid
    dearly for it: `nombres()` for sizes and calibres, `discriminants()`
    for side, colours and range words.
    """
    import par_libelle

    # Size first, isolated from the other numbers: it is what separates the
    # variants of a single range, and it alone.
    t_wms, t_tarif = taille(libelle_wms), taille(libelle_tarif)
    if t_wms and t_tarif and not (t_wms & t_tarif):
        return False

    # Range suffixes, by strict EQUALITY.
    s_wms, s_tarif = suffixes(libelle_wms), suffixes(libelle_tarif)
    if s_wms != s_tarif:
        return False

    # Letter sizes, by strict EQUALITY: S, M, L, XL separate products that
    # the rest of the label blurs together.
    l_wms, l_tarif = taille_lettre(libelle_wms), taille_lettre(libelle_tarif)
    if l_wms and l_tarif and l_wms != l_tarif:
        return False

    # Then the side: a left is not a right, at any price.
    c_wms, c_tarif = cote(libelle_wms), cote(libelle_tarif)
    if c_wms and c_tarif and c_wms != c_tarif:
        return False

    # CAPACITY, compared in millilitres: this is the discriminant that
    # really separates two pack formats of the same product.
    v_wms, v_tarif = contenance(libelle_wms), contenance(libelle_tarif)
    if not contenances_compatibles(v_wms, v_tarif):
        return False

    # Raw numbers, as a last resort ONLY.
    #
    # They cannot convert units: "MODELEX16 500 ML" carries {500} while the
    # price list "MODELEX16 NPC HF 12X0,5L" carries {12; 0.5}. Disjoint,
    # hence rejected - even though both designate the SAME 500 ml unit,
    # which the capacity check established three lines above. The coarse
    # check contradicted the fine one, and the coarse one won: MODELEX16
    # stayed unfindable even though its price was in the price list.
    #
    # Once capacity has decided on both sides, it is authoritative and we
    # ask nothing more of the raw numbers.
    if v_wms is None or v_tarif is None:
        n_wms = par_libelle.nombres(libelle_wms)
        n_tarif = par_libelle.nombres(libelle_tarif)
        if n_wms and n_tarif and not (n_wms & n_tarif):
            return False

    d_wms = par_libelle.discriminants(libelle_wms, marque)
    d_tarif = par_libelle.discriminants(libelle_tarif)
    # Present on one side only: we do not reject - the two catalogues do
    # not write the same things. Present on both sides and disjoint: these
    # are two products.
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

    # The supplier's brand distinguishes nothing: it is on both sides, on
    # every row. We remove it, along with any word present on more than half
    # the price list rows - at Supplier J, "POMPE" comes up everywhere and
    # was matching a MODELEX16 to a "POMPE 25CC CAPS ROUGE".
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
            # Same model does not mean same product: a reclining backrest
            # is not a fixed one, and they do not cost the same.
            if not traits_compatibles(mes_traits, ligne["traits"]):
                continue
            # Capacity separates two bottles of the same product.
            if not contenances_compatibles(ma_contenance, ligne["contenance"]):
                continue
            # Size and side separate two variants of a range.
            if not declinaison_compatible(art["Libellé déclinaison ^(1)"],
                                          ligne["designation"],
                                          art["Nom fournisseur"]):
                continue
            candidats.append((communs, ligne))
        if not candidats:
            continue
        # Best first: most models in common, then the price closest to the
        # DPA when there is one.
        def tri(c):
            communs, ligne = c
            prix = ligne.get("prix_achat_unitaire_ht")
            ecart = 99.0
            if dpa and prix and dpa > 0:
                r = float(prix) / float(dpa)
                ecart = abs(r - 1)
            # Purchasing rule given by the buyers: on shells, we take the
            # COMPLETE WITH TRAY model. So when the WMS says nothing, the
            # version with a tray comes first.
            tablette = ligne["traits"].get("tablette")
            sans_tablette = 1 if (tablette == "SANS"
                                  and "tablette" not in mes_traits) else 0
            return (-len(communs), sans_tablette, ecart)

        candidats.sort(key=tri)
        for rang, (communs, ligne) in enumerate(candidats[:4], 1):
            prix = ligne.get("prix_achat_unitaire_ht")
            # The price list price may cover a pack - at Supplier J,
            # "12X0,5L" is EUR 60.00 for twelve bottles (example value).
            # Without the division, the ratio to the DPA is twelve times
            # too high and the candidate
            # looks absurd when it is in fact right.
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

    # A SYSTEMATIC factor across the whole supplier - Supplier G's price
    # list quotes the pair while the WMS counts shoes one by one, and the
    # ratios then cluster around 2.00-2.05. Without this pass, correct
    # candidates fall outside the credibility band and pass for dubious.
    #
    # We reuse pa_conditionnement's inference, with its guarantee: the same
    # integer over at least five articles, and over more than 60% of them.
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
        print(f"    factor /{deduit} applied to {int(vise.sum())} rows - {preuve}")
    return res


def fournisseurs_a_traiter() -> list[str]:
    """Suppliers that have a price list AND articles without a price.

    That is exactly the population this module targets: a price list is
    there, it is read, and yet some articles are not found in it. The
    reason "aucun tarif reçu" is excluded - without a price list there is
    nothing to match, and what is needed is an email, not an algorithm.
    """
    import warnings as _w

    _w.filterwarnings("ignore")
    from extracteurs.base import chemin_lisible
    from main import SORTIE_ACHATS

    etat = SORTIE_ACHATS / "pa_etat_par_article.xlsx"
    if not etat.exists():
        raise SystemExit("Run this first: python run_pa.py")
    d = pd.read_excel(chemin_lisible(etat), skiprows=3, dtype=str)
    vises = d[d["Motif"].fillna("").str.contains(
        "pas dans son tarif", case=False)]
    comptes = vises["Fournisseur"].dropna().value_counts()
    return [str(n) for n in comptes.index]


def main() -> None:
    import mise_en_forme
    from main import SORTIE_ACHATS

    # The historical default is FOURNISSEUR_B, and it cannot be guessed:
    # running the script with no argument LOOKS LIKE a full sweep and is
    # not one. I fell for it myself on 16/09, announcing a "measurement
    # over the whole scope" that was only a pass over Supplier B. So we say
    # it out loud, and offer the real sweep.
    if "--tous" in sys.argv:
        cibles = fournisseurs_a_traiter()
        print(f"--tous: {len(cibles)} suppliers with a price list read and "
              f"articles that are not found in it")
    else:
        cibles = [a for a in sys.argv[1:] if not a.startswith("--")]
        if not cibles:
            cibles = ["FOURNISSEUR_B"]
            print("No supplier given - FOURNISSEUR_B by default.")
            print("To sweep every supplier concerned: "
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
            print(f"    {t['Code article'].nunique()} articles with at least "
                  f"one candidate, {len(t)} proposal rows")
        tables.append(t)

    utiles = [t for t in tables if not t.empty]
    if not utiles:
        print("no candidate")
        return
    tout = pd.concat(utiles, ignore_index=True)
    # Rank 1 rows with a credible ratio come to the top: that is where
    # review pays off fastest.
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
