# -*- coding: utf-8 -*-
"""
Checks EUDAMED codes obtained on a short reference against the label.

    python ean_verifier_libelle.py             resumes where it stopped
    python ean_verifier_libelle.py --perimetre the frozen scope first
    python ean_verifier_libelle.py --etat      where do we stand?
    python ean_verifier_libelle.py --essai 20  stops after 20 references

`--perimetre` brings the collection down from 19 h to 8 h by querying
only the articles of the frozen scope, the ones whose code is really
used. The cache being shared, widening the scope afterwards replays
nothing that is already done.

Output: wms_extracts/.cache_eudamed_libelle/resultats.json
        sortie/2-chantier-ean/eudamed_verification_libelle.xlsx

The problem
-----------
`eudamed_par_reference.py` queries the registry on the `reference`
parameter alone, WITHOUT a manufacturer filter, and keeps the FIRST
device returned whose primaryDi is a valid EAN. Nothing checks that this
device is ours.

Measured on 14/09: the reference "7300" (a 58 cm belt from supplier BZ)
returns 1 080 devices, the first of which is an injection tube from
supplier CA. The reference "2975" returns 909, the first of which is a
pair of Italian reading glasses. Confronted with the GS1 prefix of the
same supplier's price list, the codes obtained this way agree only
27.8 % of the time, against 100 % for a code read in a price list.

1 321 codes in V16 are in that situation, 90.8 % of what EUDAMED
returned.

The label, the only characteristic we have in common
----------------------------------------------------
We do not have the manufacturers' SRN, and the registry accepts no
filter on their name (`manufacturerName` is ignored: the answer stays
identical). `tradeName`, on the other hand, is a REAL filter, as a
substring and case-insensitive, and it combines with `reference` on the
server side.

That is what changes everything: absence becomes demonstrable. Asking
"reference=2950 AND tradeName=MODELE-B" queries the whole registry and
returns zero. That is no longer a deduction drawn from the first hundred
of nine hundred results, it is a fact.

And the same request repairs as much as it verifies. "reference=3000DBM
AND tradeName=DETACHABLE" returns two devices from supplier AH,
including "Detachable Battery Module", our own label word for word,
whereas the code kept until now came from somewhere else.

Three verdicts, never an automatic correction
---------------------------------------------
    CONFIRME   a device carrying our reference also carries our label,
               and its code is the one we already had
    CORRIGE    same thing, but the code differs: ours was wrong and this
               one is offered IN A COLUMN, not written into V16
    REJETE     no device carrying our reference carries a label close to
               ours, on any of our words
    NON CONCLU the word queried was not distinctive enough (more than a
               hundred candidates, none of them matching): to be judged
               by hand
    NON TESTE  our label offers no queryable word

This script writes NOTHING into "EAN distributeur - Vn.xlsx", which
stays the log of what the pipeline produced. It produces an opinion; it
is `ean_fiable.py` and you who draw the consequences.

Choosing the words to query
---------------------------
Our labels are French, the registry's are mostly English or German.
"CEINTURE", "CHAUSSURES", "FAUTEUIL" will never appear there: querying
them costs a minute for nothing. What does get through are the model
names, MODELE-A, MODELE-B, M300, and those alone. Hence `GENERIQUES`,
to be extended as false negatives show up.

But a hand-written list will never cover the medical vocabulary, and
every omission is paid for in minutes. The sorting is therefore done on
RARITY IN OUR OWN LABELS, measured on every pass over the 15 178
descriptions of the workbook: a word we use everywhere distinguishes
nothing, a word we use once is a model name.

    GENOUILLERE ACTIVE MODELE-A S GA T5      MODELE-A   134 uses
    CHAUSS. MODELE-B MARRON                  MODELE-B     1
    FAUTEUIL ROULANT FOURNISSEUR B M300 T42  M300        12

Sorting on length, as I did at first, put "GENOUILLERE" ahead of
"MODELE-A" and burned the first call. Length now only breaks ties at
equal frequency.

A colour word is never enough. A trial on "CHAUSS. MODELE-B MARRON"
querying MARRON brought back a Spanish shoe from supplier GA, at a score
of 0.50. That is exactly why `par_libelle`'s threshold is kept at 0.60:
a match like that falls of its own accord.

Throttling
----------
Same rule as the collection by reference: one call a minute. Below that,
EUDAMED answers 200 with zero results instead of a 429, and a failure
becomes indistinguishable from an absence, which would produce exactly
the falsehood we are trying to get out of. The state is saved after
every call: an interruption costs one request at most.
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

# A hundred candidates per call: that is the maximum the registry agrees
# to return at once, verified. Beyond that we would paginate, and
# paginating at one minute a page costs more than it brings back: if a
# hundred devices carry our reference AND a word of our label without any
# of them being ours, it is the word that was not distinctive, not the
# registry hiding the answer.
TAILLE_PAGE = 100

# Beyond three words we are querying everyday vocabulary and paying a
# minute per word. The three longest are enough: those are the model
# names.
MOTS_MAXIMUM = 3

# `par_libelle`'s threshold, taken as it is and deliberately not
# relaxed. At 0.50, a match on colour alone gets through.
SEUIL = par_libelle.SEUIL

# Commercial variants of one and the same device. `par_libelle` measures
# the share of OUR words found in the other label, never the words in
# excess, and rightly so: catalogues describe at greater length than we
# do. But a word from this list present on one side only does not
# describe: it designates a different commercial article.
#
# The first real verdict showed it. "DETACHABLE BATTERY MODULE" received,
# at a score of 1.00, "Detachable Battery Module, RENTAL", the rental
# module, not the one we sell.
#
# We do not reject it for all that: it may well be the right code, the
# registry sometimes holding a single declaration for both. We say so, in
# a column, and you are the one who decides.
VARIANTES_COMMERCIALES = {
    "RENTAL", "LOCATION", "DEMO", "DEMONSTRATION", "SAMPLE", "TRIAL",
    "REFURBISHED", "RECONDITIONNE", "SPARE", "REPLACEMENT", "RECHANGE",
    "ACCESSORY", "ACCESSOIRE", "STERILE", "NONSTERILE", "DISPOSABLE",
    "REUSABLE", "SINGLEUSE",
}

# French vocabulary that stands no chance of appearing in a registry
# label. To be extended: every word added here saves a minute per
# article that carries it.
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

# Suppliers that sent us their own code file. We do NOT query them: their
# answer holds at 98 % against the GS1 prefix test, EUDAMED at 27.8 %.
# Asking the registry what the supplier has already written means paying
# a minute for a worse answer, and risking recording a code that
# contradicts theirs.
#
# The consolidated workbook `ean_fournisseurs.xlsx` names those answers
# by FILE, not by supplier: the correspondence cannot be derived from it,
# it is written here. To be extended with every answer received.
#
# The name is the WMS one, character for character, as found in V16.
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

# How many of our labels use each word. Filled by
# `apprendre_frequences`, read by `mots_a_interroger`.
_FREQUENCES: dict[str, int] = {}


def _derniere_version() -> Path:
    """The most recent frozen version.

    Sorts on the NUMBER, never on the name: alphabetical order puts "V9"
    after "V16", and `pa_completer.py` still gets caught by that same
    rope (its glob keeps "EAN distributeur.xlsx", the working file).
    """
    fichiers = sorted(
        SORTIE_CHANTIER.glob("EAN distributeur - V*.xlsx"),
        key=lambda p: int(re.search(r"- V(\d+)", p.name).group(1)),
    )
    if not fichiers:
        raise SystemExit("No \"EAN distributeur - Vn.xlsx\" file found.")
    return fichiers[-1]


def _empecher_la_veille() -> bool:
    """Asks Windows not to fall asleep while we are working.

    Taken as it is from `eudamed_par_reference.py`: a collection that
    uses neither keyboard nor mouse counts as inactivity, and the machine
    falls asleep after a few dozen minutes. Over eight hours, it is the
    whole collection that is lost.

    We do NOT ask for ES_DISPLAY_REQUIRED: the screen can go off. The
    effect ends with the process, nothing to undo.
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
        # A domain policy may refuse it: that is no reason to stop, only
        # to say so.
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


# Words too widespread in the sector's company names to designate a
# brand. Without this filter, "LABORATOIRE FOURNISSEUR I" and "COMPRESSE
# MEDICALE" would share a word that identifies nothing. Same list as
# `pa_completer.MOTS_TROP_COURANTS`, copied rather than imported: this
# module must not depend on the PA project.
MOTS_TROP_COURANTS = {
    "CONFORT", "SANTE", "MEDICAL", "MEDICALE", "MEDICAUX", "FRANCE",
    "FRANCAISE", "SARL", "GROUPE", "LABORATOIRE", "LABORATOIRES",
    "INTERNATIONAL", "INTERNATIONALE", "DISTRIBUTION", "TECHNIQUE",
    "SERVICES", "EUROPE", "PHARMA", "SELF", "TOUS", "HYGIENE",
    "PRODUCT", "PRODUITS", "SOINS", "HEALTH", "CARE",
}

# Size equivalences between our notation and the registry's.
#
# `par_libelle`'s discriminating-word guard stops a size L from receiving
# the code of a size S, and it is indispensable. But it compares STRINGS,
# and "MEDIUM" does not look like "M":
#
#     MARQUE-U SLIP GAMME-1 MEDIUM      vs   MARQUE-U Slip Gamme-1 M 3x21p
#     MARQUE-U SLIP GAMME-2 EXTRA LARGE vs   MARQUE-U Slip Gamme-2 XL 3x21p
#
# Both matches are RIGHT and the guard rejected them, because {MEDIUM}
# and {M} differ. So we bring both sides to the same notation before
# comparing. Order matters: "EXTRA LARGE" must be seen before "LARGE",
# otherwise it becomes "EXTRA L".
TAILLES = [
    ("EXTRA EXTRA LARGE", "XXL"), ("EXTRA LARGE", "XL"),
    ("EXTRA SMALL", "XS"), ("X LARGE", "XL"), ("X SMALL", "XS"),
    ("TRES GRAND", "XL"), ("TRES PETIT", "XS"),
    ("MEDIUM", "M"), ("LARGE", "L"), ("SMALL", "S"),
    ("MOYEN", "M"), ("MOYENNE", "M"), ("GRAND", "L"), ("PETIT", "S"),
]


def harmoniser_tailles(texte) -> str:
    """Brings "MEDIUM" and "M", "EXTRA LARGE" and "XL", to one form."""
    normalise = f" {_sans_accent(texte)} "
    for longue, courte in TAILLES:
        normalise = normalise.replace(f" {longue} ", f" {courte} ")
    return normalise.strip()


def marque_probable(designation, fournisseur) -> str | None:
    """The word found both in our label and in the supplier's name.

    That is the brand, and that is what EUDAMED indexes. "HYGIENE PRODUCT
    MARQUE-U" and "MARQUE-U DISCREET MAXI" share MARQUE-U: a `reference`
    + `tradeName=MARQUE-U` request returns ONE device, the right one,
    where the reference alone returned hundreds of them.

    Sorting by rarity could not find it: MARQUE-U appears in 88 of our
    labels, so it came after GAMME-1 (9) or CHANGE (2) and was never
    submitted. The brand is frequent AT OUR END and distinctive AT
    THEIRS: the two do not contradict each other.

    With no intersection we force nothing: the first word of a label is
    not always the brand ("GENOUILLERE ACTIVE MODELE-A"), and inventing
    one would cost a minute for nothing.
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
    """How many of our labels use each word.

    Computed over ALL the labels in the workbook, not over the articles
    to be checked alone: it is our overall vocabulary that says
    "ATTELLE" is common and "MODELE-A" distinctive.
    """
    global _FREQUENCES
    compte: collections.Counter = collections.Counter()
    for designation in designations.dropna():
        compte.update({mot for mot in _sans_accent(designation).split()
                       if len(mot) >= 4 and not mot.isdigit()})
    _FREQUENCES = dict(compte)


def mots_a_interroger(designation, fournisseur=None) -> list[str]:
    """The words of our label worth a request, the rarest first.

    We strip the supplier name off the front (our labels repeat it, the
    registry does not), then everything too short, purely numeric or too
    common to distinguish anything at all.

    If `apprendre_frequences` has not been called, every word weighs zero
    and the order falls back on length: degraded, never wrong.
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

    # The brand goes AHEAD of everything else: it is the only thing our
    # label and the registry's name in the same way.
    marque = marque_probable(designation, fournisseur)
    if marque:
        retenus = [marque] + [m for m in retenus if m != marque]
    return retenus[:MOTS_MAXIMUM]


def _code_unite(primary_di) -> str | None:
    """The DI reduced to a sales unit EAN-13, or nothing.

    A GTIN-14 with an indicator from 1 to 8 designates a CASE, never the
    unit: keeping it would let a grouping code into an article
    repository. Only indicator 0 is an EAN-13 prefixed with a zero.
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
    """(candidates, announced total, status) for a reference/word pair."""
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
    # An outsized total signals an ignored filter, not a result
    if total > 100_000:
        return [], 0, "echec"
    if not total:
        return [], 0, "absent"

    # Only what serialises: the cache is JSON, and sets of words do not
    # go into it. They are rebuilt on reading, in `_verdict`: that is
    # computation, not data.
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
    """A variant word present on the registry's side only, if there is one."""
    les_notres = set(_sans_accent(notre_libelle).split())
    les_leurs = set(_sans_accent(libelle_eudamed).split())
    ecart = (les_leurs & VARIANTES_COMMERCIALES) - les_notres
    if not ecart:
        return ""
    return ("le libellé EUDAMED porte « " + ", ".join(sorted(ecart))
            + " », absent du nôtre : variante commerciale ?")


def _verdict(article: dict, etat: dict) -> dict:
    """Confronts an article with the registry, word by word, and decides.

    Stops at the first conclusive word: the following ones would cost a
    minute to confirm what we already know.
    """
    fichier = REPONSE_RECUE.get(str(article["Fournisseur"] or "").strip())
    if fichier:
        return {"verdict": "RÉPONSE FOURNISSEUR",
                "motif": f"ne pas interroger EUDAMED : {fichier}"}

    reference = str(article["Réf. fournisseur"]).strip()
    mots = mots_a_interroger(article["Désignation"], article["Fournisseur"])
    if not mots:
        return {"verdict": "NON TESTÉ", "motif": "aucun mot interrogeable"}

    # The words actually submitted. A rejection is only worth what they
    # are: "rejected after MODELE-B, CUTANE, OVALE" can be read,
    # "rejected" on its own cannot be checked. That is the first thing
    # the 14/09 trial was missing.
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

        # The sets of words are rebuilt here rather than serialised, but
        # in COPIES. Enriching them in place amounted to putting `set`
        # objects into the cache dictionaries, which are the very same
        # objects: the next save then failed, and the failure showed up
        # one row AFTER the one that had caused it.
        # Sizes are harmonised on BOTH sides before comparison, otherwise
        # "MEDIUM" and "M" disqualify each other. The original label is
        # left intact for display.
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
    """The rows belonging to the frozen scope, which we READ, never recompute.

    The key is the internal Référence. The article code only serves as a
    fallback, because a few rows of the EAN workbook have no reference,
    and losing an article on an empty field would be the worst filter of
    all.
    """
    references = perimetre_liste.references()
    codes = perimetre_liste.codes_article()
    par_reference = df.get("Référence", pd.Series("", index=df.index))
    par_reference = par_reference.fillna("").str.strip().isin(references)
    par_code = df["Code article"].fillna("").str.strip().isin(codes)
    return par_reference | par_code


def _a_verifier(perimetre_seul: bool = False) -> pd.DataFrame:
    """The rows whose code comes from EUDAMED on a short reference."""
    fichier = _derniere_version()
    df = pd.read_excel(chemin_lisible(fichier), sheet_name="Codes EAN",
                       skiprows=3, dtype=str)
    # The vocabulary is learned over the WHOLE workbook, before any
    # filter: it is our overall usage that says a word is common.
    apprendre_frequences(df["Désignation"])
    df = df[df["Exploitable"].fillna("").str.strip() == "oui"]
    douteux = df[eudamed_fiabilite.douteux(df)].copy()
    print(f"{len(douteux)} codes to check ({fichier.name})")

    if perimetre_seul:
        douteux = douteux[_dans_le_perimetre(douteux)].copy()
        print(f"  {perimetre_liste.entete()}")
        print(f"  {len(douteux)} within the frozen scope")

    print(f"  {douteux['Réf. fournisseur'].nunique()} distinct references")
    print(f"  vocabulary learned: {len(_FREQUENCES)} words")
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
    print(f"  {len(etat)} calls in cache")
    print(f"  {len(refs & interrogees)} references already started out of "
          f"{len(refs)}")
    restantes = len(refs - interrogees)
    print(f"  at least {restantes} references left, i.e. roughly "
          f"{restantes * PAUSE / 3600:.0f} h at one call a minute")


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
            print("system sleep disabled for the duration of the collection "
                  "(the screen may go off)")
        else:
            print("! sleep could not be prevented: a night of collection "
                  "may be lost if the machine falls asleep")

    etat = _charger()
    articles = _a_verifier(perimetre_seul)
    if limite:
        articles = articles.head(limite)
        print(f"  trial: {len(articles)} articles only")

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
