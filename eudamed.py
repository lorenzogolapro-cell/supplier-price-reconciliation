# -*- coding: utf-8 -*-
"""
Interrogation d'EUDAMED, la base europeenne des dispositifs medicaux.

Tout fabricant de DM mis sur le marche europeen doit y declarer ses
dispositifs avec leur UDI-DI. Or un UDI-DI EST un GTIN : c'est donc une
source publique de codes EAN, y compris pour des fournisseurs qui n'en
publient aucun dans leurs tarifs.

    EAN-13 = primaryDi (GTIN-14) prive de son zero de tete

L'API est publique et sans authentification, mais capricieuse :

  - le filtre "srn" (numero d'enregistrement du fabricant) fonctionne et
    renvoie l'integralite de son catalogue ;
  - le filtre "tradeName" fonctionne aussi, en recherche partielle et
    insensible a la casse ;
  - TOUS LES AUTRES sont ignores en silence — "manufacturerName",
    "keyword", "query"... L'API repond alors 200 avec les 3,3 millions de
    dispositifs de la base. Un total superieur a 100 000 signale donc un
    filtre non pris en compte, pas un resultat.

Il n'existe pas d'endpoint public de recherche d'acteurs : le SRN d'un
fabricant se decouvre soit par un de ses produits (via tradeName), soit
dans ses certificats CE.
"""

from __future__ import annotations

import json
import re
import time
import unicodedata
from pathlib import Path

import requests

URL = "https://ec.europa.eu/tools/eudamed/api/devices/udiDiData"
ENTETES = {
    "Accept": "application/json",
    "User-Agent": "catalogue-distributeur/1.0 (mise a jour base articles)",
}

# Au-dela, le filtre demande a ete ignore et l'API renvoie toute la base
SEUIL_FILTRE_IGNORE = 100_000

# Politesse envers un service public : une pause entre deux appels
PAUSE = 0.4
TAILLE_PAGE = 100

# SRN connus, pour eviter une decouverte a chaque execution.
# Le SRN se lit aussi dans les certificats CE publies par le fabricant.
# SRN verifies un a un : on interroge EUDAMED par nom commercial d'un
# produit connu, puis on lit la raison sociale du fabricant qui repond. La
# recherche par raison sociale, elle, ne renvoie rien — c'est une limite de
# l'API, pas une absence de declaration.
#
# Les fournisseurs P, L et AM n'y figurent pas : les recherches ne
# remontent que des homonymes (fournisseurs GC, CJ et GD). Inutile de
# reessayer.
SRN_CONNUS = {
    "FOURNISSEUR B": "BE-MF-000000001",
    "FOURNISSEUR H": "FR-MF-000000002",   # raison sociale du fournisseur H
    "FOURNISSEUR Q": "FR-MF-000000003",   # raison sociale du fournisseur Q
    "FOURNISSEUR AB": "FR-MF-000000004",  # (et non le SRN "-PR-",
                                          # qui designe un distributeur)
    "FOURNISSEUR CD": "DE-MF-000000005",
    "FOURNISSEUR BF": "FR-MF-000000006",
    "FOURNISSEUR GE": "FR-MF-000000006",  # rachete par le fournisseur BF
    "FOURNISSEUR E": "DE-MF-000000007",
}

CACHE = Path(__file__).resolve().parent / "wms_extracts" / ".cache_eudamed"


def _normaliser(texte) -> str:
    if texte is None:
        return ""
    texte = unicodedata.normalize("NFKD", str(texte))
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    return re.sub(r"[^a-z0-9]+", "", texte.lower())


def ean_depuis_gtin(primary_di) -> str | None:
    """Convertit un UDI-DI (GTIN-14) en EAN-13.

    Un GTIN-14 dont l'indicateur de conditionnement vaut 0 designe l'unite
    de vente : son EAN-13 s'obtient en retirant ce zero. Un indicateur non
    nul designe un colis, dont le code ne doit pas etre pris pour celui de
    l'unite.
    """
    if primary_di is None:
        return None
    chiffres = re.sub(r"\D", "", str(primary_di))
    if len(chiffres) == 14:
        return chiffres[1:] if chiffres[0] == "0" else None
    if len(chiffres) == 13:
        return chiffres
    return None


def _appeler(params: dict) -> tuple[int, list]:
    """Un appel a l'API. Renvoie (total, elements)."""
    reponse = requests.get(URL, params=params, headers=ENTETES, timeout=60)
    if "json" not in reponse.headers.get("content-type", ""):
        return 0, []
    donnees = reponse.json()
    return donnees.get("totalElements", 0), donnees.get("content") or []


def chercher(srn: str | None = None, trade_name: str | None = None,
             maximum: int = 1000) -> list[dict]:
    """Dispositifs correspondant a un SRN ou a un nom commercial.

    Le total est verifie : s'il depasse SEUIL_FILTRE_IGNORE, c'est que
    l'API a ignore le filtre et renvoie toute la base. On preferera ne
    rien rendre plutot que n'importe quoi.
    """
    if not srn and not trade_name:
        raise ValueError("Il faut un SRN ou un nom commercial")

    base = {"languageIso2Code": "en", "size": TAILLE_PAGE,
            "pageSize": TAILLE_PAGE}
    if srn:
        base["srn"] = srn
    if trade_name:
        base["tradeName"] = trade_name

    total, premiers = _appeler({**base, "page": 0})
    if total > SEUIL_FILTRE_IGNORE:
        # Filtre ignore : la reponse ne veut rien dire
        return []

    resultats = list(premiers)
    pages = min((total + TAILLE_PAGE - 1) // TAILLE_PAGE,
                (maximum + TAILLE_PAGE - 1) // TAILLE_PAGE)
    for page in range(1, pages):
        time.sleep(PAUSE)
        _, suite = _appeler({**base, "page": page})
        if not suite:
            break
        resultats.extend(suite)

    return resultats[:maximum]


# Mots trop generiques pour identifier une societe a eux seuls
MOTS_SOCIETE = {
    "GROUP", "GROUPE", "FRANCE", "SAS", "SARL", "SA", "GMBH", "AG", "BV",
    "NV", "LTD", "INC", "LP", "CO", "MEDICAL", "MEDIZIN", "SANTE", "HEALTH",
    "INDUSTRIES", "OPERATIONS", "SYSTEME", "SYSTEMES", "INTERNATIONAL",
}


def _mots_societe(nom) -> set:
    """Mots significatifs d'une raison sociale."""
    if not nom:
        return set()
    texte = unicodedata.normalize("NFKD", str(nom).upper())
    texte = "".join(c for c in texte if not unicodedata.combining(c))
    texte = re.sub(r"[^A-Z0-9 ]+", " ", texte)
    return {
        mot for mot in texte.split()
        if len(mot) >= 3 and mot not in MOTS_SOCIETE and not mot.isdigit()
    }


def _meme_societe(attendu, trouve) -> bool:
    """Le fabricant trouve est-il bien celui qu'on cherchait ?

    La comparaison porte sur des MOTS ENTIERS. Une simple inclusion de
    chaine ferait passer "ALPHA" pour "Alphamed Innovation", ou "BETA"
    pour "Betamedical" : on injecterait alors les codes d'un autre
    fabricant.
    """
    mots_attendus = _mots_societe(attendu)
    mots_trouves = _mots_societe(trouve)
    if not mots_attendus or not mots_trouves:
        return False
    return bool(mots_attendus & mots_trouves)


def trouver_srn(indices: list[str], fabricant: str | None = None
                ) -> tuple[str | None, str | None]:
    """Decouvre le SRN d'un fabricant a partir de noms de ses produits.

    `indices` sont des noms de modeles connus ("M200N", "Modele-D"). Le
    premier qui ramene un dispositif du bon fabricant livre son SRN.

    Renvoie (srn, nom du fabricant tel qu'EUDAMED l'ecrit).
    """

    for indice in indices:
        if not indice or len(indice) < 3:
            continue
        try:
            trouves = chercher(trade_name=indice, maximum=100)
        except requests.RequestException:
            continue
        time.sleep(PAUSE)

        for dispositif in trouves:
            nom = dispositif.get("manufacturerName") or ""
            srn = dispositif.get("manufacturerSrn")
            if not srn:
                continue
            if fabricant is None:
                return srn, nom
            # "Alpha Group" doit repondre a "ALPHA", mais "Alphamed
            # Innovation" ne doit pas repondre a "ALPHA".
            if _meme_societe(fabricant, nom):
                return srn, nom

    return None, None


def catalogue_fabricant(fournisseur: str, indices: list[str] | None = None,
                        utiliser_cache: bool = True) -> list[dict]:
    """Tous les dispositifs declares par un fabricant.

    Le SRN est cherche dans SRN_CONNUS, puis decouvert a partir des noms
    de produits fournis. Le resultat est mis en cache : la base bouge peu
    et l'API est lente.
    """
    cle = _normaliser(fournisseur)
    CACHE.mkdir(parents=True, exist_ok=True)
    fichier = CACHE / f"{cle}.json"

    if utiliser_cache and fichier.exists():
        return json.loads(fichier.read_text(encoding="utf-8"))

    srn = SRN_CONNUS.get(fournisseur.upper())
    if srn is None and indices:
        srn, nom = trouver_srn(indices, fabricant=fournisseur)
        if srn:
            print(f"    SRN découvert pour {fournisseur} : {srn} ({nom})")

    if not srn:
        # L'echec est mis en cache lui aussi : sans cela, chaque execution
        # relancerait une serie de recherches pour un fournisseur qui n'a
        # pas de SRN — un distributeur, par exemple.
        fichier.write_text("[]", encoding="utf-8")
        return []

    dispositifs = chercher(srn=srn, maximum=5000)
    fichier.write_text(
        json.dumps(dispositifs, ensure_ascii=False), encoding="utf-8"
    )
    return dispositifs


def indexer(dispositifs: list[dict]) -> dict:
    """Index des dispositifs par reference et par nom commercial.

    La reference EUDAMED porte parfois un suffixe ("REF12345 and suffix")
    qui designe une famille de declinaisons : on indexe aussi la racine.
    """
    par_reference: dict[str, dict] = {}
    par_nom: dict[str, list] = {}

    for dispositif in dispositifs:
        ean = ean_depuis_gtin(dispositif.get("primaryDi"))
        if not ean:
            continue
        enrichi = {**dispositif, "ean": ean}

        reference = dispositif.get("reference")
        if reference:
            for morceau in re.split(r"\s+and\s+|\s*[,;/]\s*", str(reference)):
                cle = _normaliser(morceau)
                if len(cle) >= 4 and cle not in par_reference:
                    par_reference[cle] = enrichi

        nom = _normaliser(dispositif.get("tradeName"))
        if nom:
            par_nom.setdefault(nom, []).append(enrichi)

    return {"par_reference": par_reference, "par_nom": par_nom}
