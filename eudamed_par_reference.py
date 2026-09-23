# -*- coding: utf-8 -*-
"""
Interroge EUDAMED reference par reference, en tache de fond.

    python eudamed_par_reference.py            reprend ou il s'etait arrete
    python eudamed_par_reference.py --etat     ou en est-on ?
    python eudamed_par_reference.py --refaire  repart de zero

Sortie : wms_extracts/.cache_eudamed_ref/resultats.json
         sortie/2-chantier-ean/ean_eudamed_reference.xlsx

Pourquoi par reference
----------------------
Toute la collecte precedente passait par le SRN du fabricant, qu'il faut
connaitre — nous n'en avons que huit. Or l'API accepte aussi un filtre
`reference`, verifie en correspondance exacte : on peut donc interroger
n'importe quel fournisseur, sans SRN, sur les references qui nous
manquent.

Le throttling, et comment on vit avec
--------------------------------------
EUDAMED limite silencieusement : passe environ trois appels par minute,
il repond HTTP 200 avec zero resultat au lieu d'un 429. Un echec est donc
indiscernable d'une absence — c'est ce qui a fige 248 caches vides qu'on
ne rejouera jamais.

Deux precautions :
  - une pause d'une minute entre deux appels, cadence a laquelle aucun
    echec n'a ete observe ;
  - un resultat vide n'est PAS considere comme definitif : la reference
    est reessayee jusqu'a REESSAIS fois, a des moments differents. Ce
    n'est qu'apres cela qu'on conclut a l'absence.

Reprise apres coupure
---------------------
L'etat est enregistre apres chaque appel. Une coupure de courant coute au
plus une reference. Le script peut etre relance autant de fois que
necessaire : il reprend la ou il en etait.
"""

from __future__ import annotations

import json
import random
import re
import sys
import time
from datetime import datetime
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import pandas as pd
import requests

import eudamed_fiabilite  # noqa: E402
import mise_en_forme  # noqa: E402
from extracteurs.base import chemin_lisible, ean_est_valide  # noqa: E402
from main import SORTIE_CHANTIER, WMS_EXTRACTS  # noqa: E402

URL = "https://ec.europa.eu/tools/eudamed/api/devices/udiDiData"
ENTETES = {
    "Accept": "application/json",
    "Accept-Language": "en",
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                   "AppleWebKit/537.36 (KHTML, like Gecko) "
                   "Chrome/120.0 Safari/537.36"),
}

CACHE = WMS_EXTRACTS / ".cache_eudamed_ref"
ETAT = CACHE / "resultats.json"

# Une minute entre deux appels : cadence a laquelle aucun echec silencieux
# n'a ete observe. En dessous, l'API repond 200 avec zero resultat.
PAUSE = 60
# On varie un peu pour ne pas taper a la seconde pres
JITTER = 12

# Combien de fois insister sur une reference restee vide.
#
# Le reglage etait a 3, par crainte du throttling silencieux. Mesure le
# 04/09 sur 1 195 references, l'hypothese ne tient pas :
#
#     1er essai   494 references   318 trouvees   64 %
#     2e essai    620 references    15 trouvees   2,4 %
#     3e essai     45 references     2 trouvees   4,4 %
#
# Un vide au premier appel est donc une VRAIE absence, pas un refus
# deguise — a une minute par appel, EUDAMED ne nous freine pas. Insister
# coutait treize heures pour une vingtaine de codes, pendant que des
# milliers de references n'avaient jamais ete interrogees une seule fois
# et avaient, elles, deux chances sur trois d'aboutir.
#
# On epuise donc les premiers essais d'abord. Un second passage se
# decidera apres, sur les absences confirmees, quand EUDAMED se sera
# rempli — la reglementation est d'application echelonnee.
REESSAIS = 1

TIMEOUT = 90


def _empecher_la_veille() -> bool:
    """Demande a Windows de ne pas s'endormir tant qu'on travaille.

    La collecte dure des jours. Le poste, lui, se met en veille au bout
    de quelques dizaines de minutes d'inactivite — et une collecte qui
    n'utilise ni clavier ni souris compte comme de l'inactivite. On perd
    alors toute la nuit, et il faut attendre l'ouverture de session du
    lendemain pour que le lanceur reprenne.

    `SetThreadExecutionState` regle cela sans aucun droit particulier :
    le processus declare que le systeme doit rester eveille. On ne
    demande PAS ES_DISPLAY_REQUIRED — l'ecran peut s'eteindre, il ne
    sert a rien ici, et le laisser allume userait la dalle pour rien.

    L'effet cesse de lui-meme quand le processus se termine : rien a
    defaire, rien qui reste accroche si la collecte est tuee.
    """
    if sys.platform != "win32":
        return False
    try:
        import ctypes
        ES_CONTINUOUS = 0x80000000
        ES_SYSTEM_REQUIRED = 0x00000001
        resultat = ctypes.windll.kernel32.SetThreadExecutionState(
            ES_CONTINUOUS | ES_SYSTEM_REQUIRED)
        return bool(resultat)
    except Exception:
        # Une politique de domaine peut le refuser : ce n'est pas une
        # raison d'interrompre la collecte, seulement de le signaler.
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


COLONNES_ARTICLE = ["Code article", "Réf. fournisseur", "Désignation",
                    "Fournisseur", "Actif"]


def _tous_les_articles() -> pd.DataFrame:
    """Tous les articles portant une reference fournisseur.

    Deux usages, et il ne faut pas les confondre : cette table sert a
    NOMMER les articles quand on ecrit le classeur des trouvailles.
    `_a_chercher` en tire ensuite ce qui reste a interroger.
    """
    fichier = mise_en_forme.derniere_version(SORTIE_CHANTIER,
                                             "EAN distributeur.xlsx")
    if fichier is None:
        raise SystemExit("Lancez d'abord : python ean_global.py")
    df = pd.read_excel(chemin_lisible(fichier), sheet_name="Codes EAN",
                       skiprows=3, dtype=str)
    return df[df["Code article"].notna() & df["Réf. fournisseur"].notna()]


# Une reference exploitable par EUDAMED : chiffres, espaces et points de
# separation. « 25 105 00 » et « 72.700 » passent, « ALPHA2_MO_T15 »
# non.
REFERENCE_NUMERIQUE = re.compile(r"[\d\s.]{4,}")


def _interrogeable(article: dict) -> bool:
    """Cette reference a-t-elle une chance d'etre connue d'EUDAMED ?

    La question ne se pose que pour les references tirees de la colonne
    « Référence fabricant » du WMS, ouverte en repli quand celle du
    fournisseur manque. Cette colonne melange de vraies references
    catalogue et des codes de configuration maison.

    Deux pilotes de 25 appels reels l'ont mesure :

        references numeriques      12 succes sur 38   32 %
        references alphanumeriques  0 succes sur 12    0 %

    « KIT_ASM_RLV01 », « GAMME MS-BASE », « ALPHA2_MO_T15 » :
    aucun fabricant ne declare un dispositif sous ce genre de chaine.
    Les interroger couterait 27 heures pour rien.

    La reference FOURNISSEUR, elle, passe sans condition : elle rend
    27,6 % quelle que soit sa forme, mesure sur 3 658 appels.

    Une condition s'y ajoute depuis le 11/09/2026, et elle vaut pour les
    DEUX colonnes : la reference doit faire au moins
    `eudamed_fiabilite.LONGUEUR_MINIMALE` caracteres. « Rendre 27,6 % »
    ne voulait pas dire « rendre juste » : la mesure comptait les appels
    qui ramenaient QUELQUE CHOSE, pas ceux qui ramenaient le bon
    dispositif. Confrontes au prefixe GS1 du fournisseur, les codes tires
    d'une reference courte tombent a 15 % de concordance, contre 96 %
    au-dela de neuf caracteres. Les interroger ne coute pas seulement du
    temps : cela produit du faux qu'il faut ensuite defaire.
    """
    reference = str(article.get("Réf. fournisseur") or "").strip()
    if not eudamed_fiabilite.reference_fiable(reference):
        return False
    if article.get("Origine référence") != "fabricant":
        return True
    return bool(REFERENCE_NUMERIQUE.fullmatch(reference))


def _a_chercher(tous: pd.DataFrame | None = None) -> list[dict]:
    """Nos manques ayant une reference exploitable, actifs d'abord."""
    df = _tous_les_articles() if tous is None else tous
    df = df[df["Code EAN"].isna()]
    # Les articles qui tournent passent en premier : si la collecte
    # s'arrete en route, c'est eux qu'on aura documentes.
    df = df.sort_values("Actif", ascending=False)

    colonnes = COLONNES_ARTICLE + (
        ["Origine référence"] if "Origine référence" in df.columns else [])
    articles = df[colonnes].to_dict("records")
    retenus = [a for a in articles if _interrogeable(a)]
    ecartes = len(articles) - len(retenus)
    if ecartes:
        print(f"{ecartes} références fabricant alphanumériques écartées "
              f"— mesuré à 0 % de rendement")
    return retenus


def _interroger(reference: str) -> tuple[str | None, str]:
    """(EAN, statut) pour une reference. statut : trouve / absent / echec."""
    try:
        reponse = requests.get(
            URL, headers=ENTETES, timeout=TIMEOUT,
            params={"languageIso2Code": "en", "size": 10, "pageSize": 10,
                    "page": 0, "reference": reference})
    except requests.RequestException:
        return None, "echec"

    if reponse.status_code != 200:
        return None, "echec"
    if "json" not in reponse.headers.get("content-type", ""):
        return None, "echec"

    donnees = reponse.json()
    total = donnees.get("totalElements", 0)
    # Un total demesure signale un filtre ignore, pas un resultat
    if total > 100_000:
        return None, "echec"
    if not total:
        return None, "absent"

    for element in donnees.get("content") or []:
        brut = str(element.get("primaryDi") or "")
        if len(brut) == 14 and brut[0] == "0":
            candidat = brut[1:]
        elif len(brut) == 13:
            candidat = brut
        else:
            continue
        if ean_est_valide(candidat):
            return candidat, "trouve"
    # Des dispositifs existent mais aucun code d'unite de vente
    return None, "absent"


def _ecrire_resultats(etat: dict, tous: pd.DataFrame) -> None:
    """Classeur des codes trouves, pret a etre repris.

    Ecrit d'apres TOUS les articles, jamais d'apres la liste de ce qui
    reste a chercher. La nuance a coute 221 codes : cette liste ne
    contient que les articles SANS code, or un article qu'EUDAMED vient
    de documenter n'en fait plus partie au passage suivant. Son code
    disparaissait donc du classeur, `ean_global` ne le voyait plus, et
    l'article se retrouvait de nouveau sans code — puis de nouveau dans
    la file. La collecte effacait ses propres trouvailles.

    Une reference peut nommer plusieurs de nos articles : le code vaut
    alors pour chacun d'eux, et chacun recoit sa ligne.
    """
    par_reference: dict[str, list[dict]] = {}
    for article in tous[COLONNES_ARTICLE].to_dict("records"):
        par_reference.setdefault(article["Réf. fournisseur"], []).append(
            article)

    lignes = []
    for reference, valeur in etat.items():
        if not valeur.get("ean"):
            continue
        for article in par_reference.get(reference, []):
            lignes.append({
                "Code article": article["Code article"],
                "Référence fournisseur": reference,
                "Désignation": article["Désignation"],
                "Fournisseur": article["Fournisseur"],
                "Actif": article["Actif"],
                "Code EAN": valeur["ean"],
                "Source": "eudamed-référence",
                "Relevé le": valeur.get("date", ""),
            })
    if not lignes:
        return

    df = pd.DataFrame(lignes)
    SORTIE_CHANTIER.mkdir(parents=True, exist_ok=True)
    chemin = mise_en_forme.chemin_ecriture(
        SORTIE_CHANTIER / "ean_eudamed_reference.xlsx")
    with pd.ExcelWriter(chemin, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Codes EAN", index=False, startrow=3)
    mise_en_forme.formater(
        chemin,
        options={"Codes EAN": {
            "ligne_entete": 4,
            "titre": "Codes EAN relevés sur EUDAMED, par référence",
            "sous_titre": (
                f"{len(df)} codes | correspondance exacte sur la référence "
                f"fabricant, donc sûre"
            ),
        }},
    )


def etat_courant() -> None:
    etat = _charger()
    articles = _a_chercher(_tous_les_articles())
    trouves = sum(1 for v in etat.values() if v.get("ean"))
    absents = sum(1 for v in etat.values()
                  if v.get("statut") == "absent"
                  and v.get("essais", 0) >= REESSAIS)
    a_reessayer = sum(1 for v in etat.values()
                      if not v.get("ean") and v.get("essais", 0) < REESSAIS)
    restants = [a for a in articles
                if etat.get(a["Réf. fournisseur"], {}).get("essais", 0)
                < REESSAIS and not etat.get(a["Réf. fournisseur"],
                                            {}).get("ean")]
    print(f"  {len(articles)} références à documenter")
    print(f"  {trouves} trouvées | {absents} absentes (confirmé) | "
          f"{a_reessayer} à réessayer")
    print(f"  reste {len(restants)} appels, soit environ "
          f"{len(restants) * PAUSE / 3600:.0f} h")


def main() -> None:
    if "--refaire" in sys.argv and ETAT.exists():
        ETAT.unlink()
    if "--etat" in sys.argv:
        etat_courant()
        return

    if _empecher_la_veille():
        print("veille système désactivée le temps de la collecte "
              "(l'écran peut s'éteindre)")
    else:
        print("! la mise en veille n'a pas pu être empêchée : une nuit de "
              "collecte peut être perdue si le poste s'endort")

    etat = _charger()
    tous = _tous_les_articles()
    articles = _a_chercher(tous)

    # Ce qui reste : jamais tente, ou tente sans succes moins de REESSAIS.
    #
    # UNE SEULE FOIS PAR REFERENCE. Plusieurs de nos articles partagent
    # souvent la meme reference fournisseur — declinaisons creees
    # separement, ou doublons de la base. Sans ce dedoublonnage, chacun
    # declenchait son propre appel pour la meme question : une reference
    # a ete interrogee 29 fois, une autre 10, et le compteur d'essais
    # grimpait sans que rien de nouveau soit demande. C'est du temps
    # d'appel perdu, et c'est aussi ce qui faisait croire a un
    # acharnement sur des references epuisees.
    restants, vues = [], set()
    for article in articles:
        reference = article["Réf. fournisseur"]
        if reference in vues:
            continue
        vues.add(reference)
        connu = etat.get(reference, {})
        if connu.get("ean") or connu.get("essais", 0) >= REESSAIS:
            continue
        restants.append(article)

    trouves = sum(1 for v in etat.values() if v.get("ean"))
    print(f"{len(articles)} articles | {len(vues)} références distinctes | "
          f"{trouves} déjà trouvées")
    print(f"{len(restants)} à interroger, ~{len(restants) * PAUSE / 3600:.0f} h "
          f"à raison d'un appel par minute\n")

    depuis_ecriture = 0
    for rang, article in enumerate(restants, start=1):
        reference = article["Réf. fournisseur"]
        ean, statut = _interroger(reference)

        entree = etat.setdefault(reference, {"essais": 0})
        entree["essais"] = entree.get("essais", 0) + 1
        entree["statut"] = statut
        entree["date"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        if ean:
            entree["ean"] = ean
            trouves += 1

        marque = {"trouve": "OK  ", "absent": "  --", "echec": " !! "}[statut]
        print(f"  [{rang:>5}/{len(restants)}] {marque} "
              f"{str(reference)[:20]:<20} {ean or ''}  "
              f"({trouves} trouvés)", flush=True)

        _enregistrer(etat)
        depuis_ecriture += 1
        if depuis_ecriture >= 25:
            _ecrire_resultats(etat, tous)
            depuis_ecriture = 0

        if rang < len(restants):
            time.sleep(PAUSE + random.uniform(0, JITTER))

    _ecrire_resultats(etat, tous)
    print(f"\n{trouves} codes EAN trouvés sur EUDAMED")


if __name__ == "__main__":
    main()
