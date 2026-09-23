# -*- coding: utf-8 -*-
"""
Chantier PA — ce qu'on met de côté, en deux listes et rien d'autre.

    python pa_ecartes.py

POURQUOI TROIS FEUILLES, ET PAS UNE
    Ce sont trois décisions de nature différente, et les confondre ferait
    perdre l'information :

    « Écarté »            le produit existe et se vend ; c'est NOUS qui le
                          remettons à plus tard — lits, fauteuils roulants,
                          releveurs, VPH. Décision de périmètre, réversible.

    « Fournisseur arrêté » on n'achète plus chez EUX ; le produit lui-même
                          n'a rien. Constat des appros, daté
                          (`fournisseurs_arretes.py`), pas une supposition.

    « Arrêt fabricant »   le produit est introuvable au catalogue du
                          fournisseur ; il est probablement arrêté. Constat
                          sur le monde extérieur, pas une décision interne.

    Un article écarté reviendra peut-être ; un fournisseur arrêté ou un
    produit arrêté, non. D'où trois feuilles.

CE QUI REMPLIT CHAQUE FEUILLE
    Écarté             l'annotation portée à la main dans pa_decisions.xlsx,
                       PLUS toute la famille lourde du périmètre (FAMILLES),
                       PLUS les fournisseurs mis de côté par la direction.
    Fournisseur arrêté `fournisseurs_arretes.est_arrete()`, regardé sur le
                       Fournisseur ET le fabricant — sinon un article dont
                       le WMS ne renseigne QUE le fabricant échappe au
                       filtre. C'est arrivé le 18/09 : le 48553 (Fournisseur
                       T, champ Fournisseur vide, fabricant « FRST FRANCE
                       SAS ») avait reçu un vrai prix 2026 malgré l'arrêt
                       des commandes constaté le 15/09.
    Arrêt fabricant    l'annotation à la main, et elle SEULE. Rien ici
                       n'est déduit : « arrêté » est une affirmation sur un
                       produit, elle ne se devine pas.

    Le WMS ne date ni ne signale l'arrêt d'un PRODUIT (aucun Type 79 au
    périmètre au 17/09) : il n'existe donc aucune corroboration
    automatique pour « Arrêt fabricant ». La colonne « Votre annotation »
    est la preuve, et la seule. L'arrêt d'un FOURNISSEUR, lui, est écrit
    et daté ailleurs — ce n'est pas la même source, ni le même niveau de
    certitude.

LE FOURNISSEUR MANQUANT
    423 articles du périmètre n'ont pas de fournisseur. 420 ont un
    fabricant dans l'extract WMS, et c'est lui qu'on reprend — la colonne
    « Origine du fournisseur » dit toujours d'où vient la valeur, pour
    qu'on ne confonde jamais un fournisseur su et un fabricant supposé.
    Les 3 restants sont nommés dans la sortie console.

RIEN N'EST ÉCRIT AILLEURS
    Ce module lit et range. Il ne touche ni au classeur du responsable des
    prix, ni aux prix, ni au périmètre.
"""

from __future__ import annotations

import re
import sys
import unicodedata
from datetime import date
from functools import lru_cache
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import warnings  # noqa: E402

warnings.filterwarnings("ignore")

import pandas as pd  # noqa: E402

import fournisseurs_arretes  # noqa: E402
import mise_en_forme  # noqa: E402
import pa_annotations_livrable  # noqa: E402
from extracteurs.base import chemin_lisible  # noqa: E402
from main import SORTIE_ACHATS  # noqa: E402

ETAT = SORTIE_ACHATS / "pa_etat_par_article.xlsx"
DECISIONS = SORTIE_ACHATS / "pa_decisions.xlsx"
SORTIE = SORTIE_ACHATS / "pa_ecartes_et_arrets.xlsx"

# Les familles lourdes, au mot près du WMS. Décision de la direction le
# 17/09 : lits, fauteuils roulants, releveurs et VPH sortent du chantier
# courant — trop d'options, trop de déclinaisons, un tarif par
# configuration. On les range, on ne les jette pas.
#
# « Transfert » n'y est PAS : les guidons de transfert se cotent
# normalement et on en a déjà rapproché (cf. l'article 31128).
FAMILLES = (
    "Fauteuil Roulant",
    "FAUTEUIL ROULANT MANUEL",
    "Accessoires pour fauteuils roulants",
    "Faut releveurs et repos",
    "Lit",
    "Scooters  tricycles poussettes",
    "Elevateur de bain",
    "LES APPAREILS DE VERTICALISATION",
)

# Mis de côté par la direction. On regarde le FABRICANT autant que le
# fournisseur : 67 articles portent le Fournisseur H en fabricant sans le
# porter en fournisseur, et passaient au travers du filtre.
MIS_DE_COTE = re.compile(r"FRSB|FRSH|FRSR", re.I)

# Ce que veulent dire vos mots. L'arrêt l'emporte : il est plus précis
# qu'un écartement de famille, et il vient d'un constat au catalogue.
#
# « inexistant » s'écrit à la main, donc rarement deux fois pareil : on a
# lu « innexistant » ET « inenxistant » le même jour. D'où le trou de
# trois lettres entre « in » et « xist » — une annotation perdue pour une
# coquille, c'est un article qu'on croit encore vendable.
DIT_ARRET = re.compile(
    r"arr[eêé]t|n.existe plus|\bin[a-z]{0,3}xist|introuvable", re.I)
DIT_ECARTE = re.compile(r"[eé]cart|releveur|\bvph\b|exclu", re.I)

# SANS CATALOGUE — le fournisseur ne nous a jamais envoyé de tarif (18/09)
#
# Quatrième catégorie, et elle ne dit pas la même chose que les trois
# autres. « Écarté » est un choix, « arrêté » est un constat sur le
# produit ou la relation ; ici il n'y a NI choix NI arrêt : il manque
# simplement la matière première. Le produit se vend, le fournisseur
# existe, on n'a rien à lire.
#
# Confondre les deux fausse la lecture dans les deux sens : ça gonfle le
# reste-à-faire d'articles sur lesquels aucun travail n'est possible, et
# ça masque le seul geste qui les débloquerait — réclamer le tarif.
#
# COMMENT LA CATÉGORIE SE REMPLIT — deux voies, et la première suffit
# presque toujours :
#
#   1. DÉRIVÉE du motif « aucun tarif reçu de ce fournisseur », que
#      `pa_etat_par_article` pose déjà. 56 fournisseurs, 313 articles
#      actifs au 18/09. Aucune liste à tenir : un fournisseur qui envoie
#      enfin son tarif sort de la catégorie tout seul, à la passe
#      suivante.
#
#   2. NOMMÉE à la main ci-dessous, pour ce que le motif ne peut pas
#      attraper — au premier chef les articles SANS fournisseur dans le WMS,
#      qui portent un autre motif (« aucun fournisseur au périmètre ») et
#      dont seul le fabricant est connu : le Fournisseur FA en est le cas
#      type, ses 12 articles n'ont pas de fournisseur du tout.
#
# Liste dictée par le métier le 18/09. On matche sur le fournisseur ET sur
# le fabricant, pour la raison ci-dessus.
#
# NE SONT PAS DANS CETTE LISTE, volontairement :
#   Les Fournisseurs Q et Q2 — on a leur catalogue
#       (`catalogues/FRSQ-FRSQ2/2026/`), ce sont deux noms d'un même
#       fournisseur. Si leurs articles ressortent sans prix, c'est un défaut
#       de rapprochement, pas un tarif manquant — voir la note dans
#       `classer()`.
#   Le Fournisseur FP — repris par le Fournisseur H : c'est un fournisseur
#       ARRÊTÉ, il n'y aura jamais de catalogue
#       (`fournisseurs_arretes.ARRETES`).
SANS_CATALOGUE = (
    "FRSFA",
    "FRSFB MEDICAL",
    "FRSFC",
    "FRSFD",
    "FRSFE",
    "FRSFF",
    "FRSBS",
    "FRSFG",
    "FRSBN",
    "FRSFH",
    "FRSFI",              # « SAS FRSFI - FRSFI BIS » : les deux raisons
    "FRSFI BIS",          #   sociales designent le même fournisseur
    "FRSFJ CENTRE",       # le sigle seul attraperait deux autres
                          #   fournisseurs dont le nom le contient
    "FRSFK",
    "FRSFL",
    "FRSFM NORD",         # le sigle seul est trop court pour un mot entier
    "FRSFN-SUFFIXE",
    "FRSFN SUFFIXE",      # deux orthographes, trait d'union ou espace
    "FRSFO",
)

_SANS_CATALOGUE = re.compile(
    "|".join(re.escape(nom) for nom in SANS_CATALOGUE), re.I)

# MARQUES QUI IDENTIFIENT UN FOURNISSEUR DANS UN LIBELLÉ (18/09)
#
# Règle posée par le métier : « si y'a le nom d'un fournisseur dans le
# libellé, tu mets dans arrêt ». Le cas qui l'a motivée :
#
#   46675  STETHOSCOPE ... MODELE FRSBD GRIS  -> Fournisseur A,  sans prix
#   46677  STETHOSCOPE ... MODELE FRSBD VERT  -> Fournisseur BD, a un prix
#
# Deux fois le même produit : celui qui est correctement rattaché est
# chiffré, l'autre non. L'article n'est pas introuvable, il est MAL
# RATTACHÉ dans le WMS — et on ne corrige pas le WMS depuis ce dépôt.
#
# LISTE CURÉE, PAS DÉRIVÉE — et c'est délibéré. Dériver les marques des
# 127 noms de fournisseurs déclenchait 14 fois, dont QUATRE sur le mot
# « CONFORT » : « OPTION ASSISE CONFORT ROLLATOR » n'a rien à voir avec
# la raison sociale du Fournisseur CT ni avec celle du Fournisseur BV, qui
# contiennent toutes deux ce mot. C'est le même piège que
# `MOTS_TROP_COURANTS` ailleurs dans le chantier : un mot de français
# courant ne peut pas servir d'identifiant.
#
# N'ajouter ici qu'un nom qui ne veut RIEN DIRE d'autre qu'une marque.
MARQUES_FOURNISSEUR = (
    "FRSBD", "FRSV", "FRSF", "FRSG", "FRSC", "FRSAU",
    "FRSAE", "FRSAK", "FRSAF", "FRSD", "FRSX", "FRSO",
    "FRSE", "FRSB", "FRSH", "FRSL", "FRSBF",
    "FRSBG", "FRSFQ", "FRSW", "FRSAI",
)


def _plein(serie: pd.Series) -> pd.Series:
    return serie.fillna("").astype(str).str.strip().replace("nan", "")


def _aplati(texte) -> str:
    """Un nom réduit à ses lettres et chiffres, pour comparer sans se faire
    piéger par la ponctuation (« FRS.AE MEDICAL » vs « FRSAE »)."""
    sans_accent = unicodedata.normalize("NFKD", str(texte or "").upper())
    sans_accent = "".join(c for c in sans_accent
                          if not unicodedata.combining(c))
    return re.sub(r"[^A-Z0-9]", "", sans_accent)


@lru_cache(maxsize=1)
def _catalogues_connus() -> tuple:
    """Les noms sous lesquels un catalogue existe : dossiers + registre.

    Sert à répondre à UNE question : « a-t-on quelque chose à lire pour ce
    fournisseur ? ». Deux sources parce qu'elles ne se recouvrent pas — un
    dossier peut exister sans entrée au registre (personne ne le lit
    encore), et un motif du registre peut viser un dossier au nom
    différent.
    """
    from main import CATALOGUES, FOURNISSEURS

    noms = set()
    if CATALOGUES.is_dir():
        for dossier in CATALOGUES.iterdir():
            if dossier.is_dir() and dossier.name != "Archives":
                # « FRSQ-FRSQ2 » vaut pour FRSQ ET pour FRSQ2 : on éclate
                # sur les séparateurs, sinon FRSQ2 passerait pour sans
                # catalogue alors que son tarif est dans ce dossier.
                for morceau in re.split(r"[-+&/]", dossier.name):
                    if len(_aplati(morceau)) >= 4:
                        noms.add(_aplati(morceau))
    for entree in FOURNISSEURS:
        if len(_aplati(entree["motif_wms"])) >= 4:
            noms.add(_aplati(entree["motif_wms"]))
    return tuple(sorted(noms))


def a_un_catalogue(nom) -> bool:
    """Existe-t-il un catalogue pour ce fournisseur ou ce fabricant ?"""
    p = _aplati(nom)
    if len(p) < 4:
        return False
    return any(c in p or p in c for c in _catalogues_connus())


def perimetre() -> pd.DataFrame:
    d = pd.read_excel(chemin_lisible(ETAT), skiprows=3, dtype=str)
    for c in d.columns:
        d[c] = _plein(d[c])
    d["Code article"] = d["Code article"].str.replace(r"\.0$", "", regex=True)
    d["Vendu 12 mois"] = d["VENTE_12M"].str.lower().map(
        {"true": "oui"}).fillna("")

    # Le fournisseur manquant se comble par le fabricant, et ça se dit.
    manque = d["Fournisseur"] == ""
    d["Origine du fournisseur"] = "WMS"
    d.loc[manque, "Origine du fournisseur"] = "introuvable"
    repris = manque & (d["Nom fabriquant"] != "")
    d.loc[repris, "Fournisseur"] = d.loc[repris, "Nom fabriquant"]
    d.loc[repris, "Origine du fournisseur"] = "fabricant (extract WMS)"
    return d


def annotations() -> pd.DataFrame:
    """Vos mots, repris tels quels, avec la feuille d'où ils viennent."""
    if not DECISIONS.exists():
        return pd.DataFrame(columns=["Code article", "Votre annotation",
                                     "Feuille d'origine"])
    x = pd.ExcelFile(chemin_lisible(DECISIONS))
    morceaux = []
    for feuille in x.sheet_names:
        if feuille.startswith(("0", "1")):
            continue  # la feuille 1 décide par fournisseur, pas par article
        d = pd.read_excel(x, sheet_name=feuille, skiprows=3, dtype=str)
        if "Décision" not in d.columns or "Code article" not in d.columns:
            continue
        d = d[["Code article", "Décision"]].copy()
        d["Code article"] = _plein(d["Code article"]).str.replace(
            r"\.0$", "", regex=True)
        d["Votre annotation"] = (_plein(d["Décision"])
                                 .str.replace(r"\s+", " ", regex=True))
        d = d[d["Votre annotation"] != ""]
        d["Feuille d'origine"] = feuille
        morceaux.append(d[["Code article", "Votre annotation",
                           "Feuille d'origine"]])
    if not morceaux:
        return pd.DataFrame(columns=["Code article", "Votre annotation",
                                     "Feuille d'origine"])
    a = pd.concat(morceaux, ignore_index=True)
    return a.drop_duplicates(subset=["Code article"], keep="first")


COLONNES = ["Fournisseur", "Origine du fournisseur", "Code article",
            "Réf. interne", "Désignation", "Famille", "Statut PA",
            "Vendu 12 mois", "Ancien PA", "Nouveau PA",
            "Réf. article fournisseur", "Référence fabricant", "Motif",
            "Votre annotation", "Feuille d'origine"]


def _mettre_en_colonnes(d: pd.DataFrame, motif: pd.Series) -> pd.DataFrame:
    t = d.copy()
    t["Motif"] = motif
    colonnes = [c for c in COLONNES if c in t.columns]
    return t[colonnes]


def classer(d: pd.DataFrame) -> dict:
    """Le classement, ligne par ligne — LA fonction de référence.

    Elle travaille sur n'importe quelle table portant `Famille`,
    `Fournisseur`, `Nom fabriquant` et, si elle existe, `Votre annotation`.
    Elle ne lit aucun fichier : c'est ce qui permet à `pa_etat_par_article`
    de l'appeler PENDANT qu'il construit le socle, alors que le fichier
    dont `pa_ecartes` part d'habitude n'existe pas encore.

    POURQUOI UNE SEULE FONCTION, ET PAS DEUX COPIES
        Le socle et ce module doivent classer IDENTIQUEMENT, sinon la
        colonne « Mise à l'écart » du socle et les feuilles produites ici
        se contrediront un jour — et personne ne saura laquelle croit. Une
        règle écrite deux fois est une règle qui divergera.

    Rend quatre masques booléens, plus le motif en clair.
    """
    dit = (_plein(d["Votre annotation"]) if "Votre annotation" in d.columns
           else pd.Series("", index=d.index))
    arret = dit.str.contains(DIT_ARRET) & ~dit.str.contains(DIT_ECARTE)
    ecarte_dit = dit.str.contains(DIT_ECARTE) & ~arret

    # Les annotations portées dans le LIVRABLE, relevées au registre
    # durable (`pa_annotations_livrable`). Elles s'ajoutent à celles de
    # `pa_decisions.xlsx` : ce sont deux endroits où la même personne
    # écrit la même sorte de décision, à deux moments différents.
    codes = _plein(d["Code article"])
    arret |= codes.isin(pa_annotations_livrable.par_lecture("arrêt fabricant"))
    annote_sans_catalogue = codes.isin(
        pa_annotations_livrable.par_lecture("sans catalogue"))

    # LE NOM D'UN AUTRE FOURNISSEUR DANS LE LIBELLÉ (règle du 18/09)
    #
    # « STETHOSCOPE MODELE FRSBD GRIS » est rattaché au Fournisseur A
    # dans le WMS alors que son libellé dit FRSBD — et son jumeau vert,
    # correctement rattaché au Fournisseur BD, a bien un prix. L'article
    # n'est pas introuvable : il est mal rattaché, et on ne corrige pas le
    # WMS depuis ici.
    #
    # On ne retient que les fournisseurs RÉELS du périmètre, pas n'importe
    # quel mot : un libellé qui contient un nom commun homonyme d'une
    # marque produirait sinon des faux positifs en série.
    # DEUX GARDE-FOUS, chacun payé par un faux positif mesuré le 18/09 :
    #
    #   1. On compare des noms APLATIS (sans point ni espace). Sans cela,
    #      « FRSAE <produit> » chez « FRS.AE MEDICAL » passait pour un
    #      article d'un AUTRE fournisseur, à cause du seul point : 21
    #      articles du Fournisseur AE, tous déjà chiffrés, partaient en
    #      « arrêt ».
    #
    #   2. La règle ne vaut QUE pour les articles SANS PRIX. Un article
    #      chiffré n'a pas de problème de rattachement à régler — quoi que
    #      dise son libellé, le prix est là et il est juste.
    autre_fournisseur = pd.Series(False, index=d.index)
    if "Désignation" in d.columns:
        libelle = _plein(d["Désignation"]).str.upper()
        aplati = (_plein(d["Fournisseur"]).str.upper()
                  .str.replace(r"[^A-Z0-9]", "", regex=True))
        for marque in MARQUES_FOURNISSEUR:
            porte = libelle.str.contains(rf"\b{marque}\b", regex=True)
            pas_le_sien = ~aplati.str.contains(marque, regex=False)
            autre_fournisseur |= porte & pas_le_sien

    if "Nouveau PA" in d.columns:
        sans_prix = pd.to_numeric(d["Nouveau PA"], errors="coerce").isna()
        autre_fournisseur &= sans_prix

    arret |= autre_fournisseur

    # « SI ILS ONT UN CATALOGUE, ILS ONT UN PRIX » — règle du métier, 18/09
    #
    # C'est la règle qui ferme la boucle, et elle était déjà écrite dans
    # les données sans qu'on l'exploite. `pa_etat_par_article` distingue
    # depuis toujours deux motifs d'absence de prix :
    #
    #   « fournisseur tarifé, mais l'article n'est pas dans son tarif »
    #        -> on A le catalogue, la référence n'y figure pas.
    #           Donc le produit est sorti de la gamme : ARRÊT.
    #
    #   « aucun tarif reçu de ce fournisseur »
    #        -> on n'a RIEN à lire : SANS CATALOGUE (traité plus bas).
    #
    # SUR QUOI ÇA REPOSE : le constat est direct — LE CATALOGUE DU
    # FOURNISSEUR NE PROPOSE PAS CE PRODUIT. On a son tarif, on l'a lu en
    # entier, la référence n'y est pas. Le programme ne devine rien : il
    # rapproche ce qu'il a lu et constate une absence.
    #
    # La réserve est étroite et porte sur des cas isolés : une référence
    # qui aurait changé côté fournisseur produirait la même absence. C'est
    # pourquoi le classement se REFAIT à chaque passe et n'est jamais
    # figé — tarif complété ou référence corrigée dans le WMS, et l'article
    # revient de lui-même au chantier parce que son motif aura changé.
    #
    # Le motif n'est rempli QUE sur les articles sans prix (règle de
    # `pa_etat_par_article`) : s'y appuyer restreint donc d'office la
    # portée, sans avoir à le redire.
    if "Motif" in d.columns:
        motif_lu = _plein(d["Motif"])
        arret |= motif_lu.str.contains("n'est pas dans son tarif",
                                       regex=False)

        # « Aucun fournisseur au périmètre — voir le fabricant » : la fiche
        # du WMS ne dit pas chez qui on achète, seul le fabricant est connu.
        # C'est donc LUI qui répond à la question « a-t-on un catalogue ? »,
        # et la même règle s'applique ensuite : catalogue + pas de prix =
        # arrêt ; pas de catalogue = sans catalogue (plus bas).
        via_fabricant = motif_lu.str.contains("voir le fabricant",
                                              regex=False)
        if via_fabricant.any():
            fabricant_servi = _plein(d["Nom fabriquant"]).map(a_un_catalogue)
            arret |= via_fabricant & fabricant_servi

    lourd = _plein(d["Famille"]).isin(FAMILLES)
    parque = (_plein(d["Fournisseur"]).str.contains(MIS_DE_COTE)
              | _plein(d["Nom fabriquant"]).str.contains(MIS_DE_COTE))

    # Fournisseur arrêté (Fournisseurs T, BC…) : regardé sur le Fournisseur
    # ET le fabricant. Le statut « 4 — HORS CHANTIER » le dit déjà pour la
    # plupart, mais seulement quand le WMS renseigne « Fournisseur » — un
    # article où seul le fabricant est connu (48553) lui échappe.
    fournisseur_arrete = (
        _plein(d["Fournisseur"]).map(fournisseurs_arretes.est_arrete)
        | _plein(d["Nom fabriquant"]).map(fournisseurs_arretes.est_arrete))
    if "Statut PA" in d.columns:
        fournisseur_arrete |= _plein(d["Statut PA"]).str.startswith("4")

    # SANS CATALOGUE : dérivé du motif quand il existe, nommé sinon.
    #
    # Les Fournisseurs Q et Q2 sont exclus explicitement : leur catalogue
    # existe (`FRSQ-FRSQ2/2026`). Si un de leurs articles ressort sans prix,
    # c'est que le rapprochement a échoué — un tout autre problème, qu'il
    # ne faut pas masquer en le rangeant ici. Au 18/09 le registre ne
    # cherche d'ailleurs que « FRSQ » : les articles FRSQ2 ne sont jamais
    # confrontés à ce tarif, ce qui est un défaut à corriger, pas une
    # absence de source.
    a_un_catalogue_q = (
        _plein(d["Fournisseur"]).str.contains(r"FRSQ|FRSQ2", regex=True)
        | _plein(d["Nom fabriquant"]).str.contains(r"FRSQ|FRSQ2", regex=True))
    sans_catalogue = (_plein(d["Fournisseur"]).str.contains(_SANS_CATALOGUE)
                      | _plein(d["Nom fabriquant"]).str.contains(_SANS_CATALOGUE))
    if "Motif" in d.columns:
        motif_lu = _plein(d["Motif"])
        sans_catalogue |= motif_lu.str.contains("aucun tarif reçu",
                                                regex=False)
        # Sans fournisseur dans le WMS ET sans catalogue chez le fabricant :
        # il n'y a rien à lire nulle part.
        sans_catalogue |= (motif_lu.str.contains("voir le fabricant",
                                                 regex=False)
                           & ~_plein(d["Nom fabriquant"]).map(a_un_catalogue))
    sans_catalogue |= annote_sans_catalogue
    sans_catalogue &= ~a_un_catalogue_q

    # Le motif, dans l'ordre où il prime. La première affectation est la
    # plus faible, la dernière l'emporte : un fauteuil roulant acheté chez
    # un fournisseur sans catalogue reste d'abord un fauteuil roulant —
    # réclamer son tarif ne servirait à rien tant qu'il est mis de côté.
    motif = pd.Series("", index=d.index)
    motif[sans_catalogue] = ("sans catalogue — aucun tarif reçu de "
                             + _plein(d["Fournisseur"])[sans_catalogue]
                             .replace("", "ce fournisseur").str.slice(0, 40))
    motif[parque] = ("fournisseur mis de côté par la direction : "
                     + _plein(d["Fournisseur"])[parque].str.slice(0, 40))
    motif[lourd] = "famille mise de côté : " + _plein(d["Famille"])[lourd]
    motif[ecarte_dit] = "votre annotation"
    motif[arret] = "arrêt fabricant — introuvable au catalogue"
    motif[fournisseur_arrete] = (
        _plein(d["Fournisseur"])[fournisseur_arrete]
        .map(fournisseurs_arretes.motif)
        .where(lambda s: s != "",
              _plein(d["Nom fabriquant"])[fournisseur_arrete]
              .map(fournisseurs_arretes.motif))
        .where(lambda s: s != "", "fournisseur arrêté"))

    return {"arret": arret, "ecarte_dit": ecarte_dit, "lourd": lourd,
            "parque": parque, "fournisseur_arrete": fournisseur_arrete,
            "sans_catalogue": sans_catalogue, "motif": motif}


def construire() -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    d = perimetre()
    a = annotations()
    d = d.merge(a, on="Code article", how="left")
    d["Votre annotation"] = _plein(d["Votre annotation"])
    d["Feuille d'origine"] = _plein(d["Feuille d'origine"])

    c = classer(d)
    arret = c["arret"]
    ecarte_dit = c["ecarte_dit"]
    lourd = c["lourd"]
    parque = c["parque"]
    fournisseur_arrete = c["fournisseur_arrete"]
    sans_catalogue = c["sans_catalogue"]
    motif = c["motif"]

    # CHAQUE ARTICLE DANS UNE SEULE FEUILLE, celle de son motif retenu.
    #
    # Sans cette exclusion mutuelle, les totaux des quatre feuilles ne
    # s'additionnent plus : mesuré le 18/09, elles affichaient 2 225
    # lignes pour 1 780 articles réellement mis à l'écart — 445 comptés
    # deux fois, parce qu'un fauteuil arrêté figurait à la fois en
    # « Écarté » et en « Arrêt fabricant ». Un lecteur qui additionne les
    # onglets tombe alors sur un chiffre qui n'existe pas.
    #
    # L'ordre est CELUI DU MOTIF, pour que la feuille où l'article se
    # trouve corresponde toujours à l'étiquette qu'il porte. Du fait le
    # plus dur au choix le plus révisable :
    #     fournisseur arrêté > arrêt fabricant > écarté > sans catalogue
    arret = arret & ~fournisseur_arrete
    a_ecarter = (ecarte_dit | lourd | parque) & ~(fournisseur_arrete | arret)
    sans_catalogue = sans_catalogue & ~(fournisseur_arrete | arret
                                        | a_ecarter)
    feuille_ecarte = _mettre_en_colonnes(d[a_ecarter], motif[a_ecarter])
    feuille_ecarte = feuille_ecarte.sort_values(
        ["Statut PA", "Fournisseur", "Désignation"],
        ascending=[False, True, True])

    feuille_fournisseur_arrete = _mettre_en_colonnes(
        d[fournisseur_arrete], motif[fournisseur_arrete])
    feuille_fournisseur_arrete = feuille_fournisseur_arrete.sort_values(
        ["Fournisseur", "Désignation"])

    feuille_arret = _mettre_en_colonnes(
        d[arret], pd.Series("introuvable au catalogue fournisseur",
                            index=d.index)[arret])
    feuille_arret = feuille_arret.sort_values(["Fournisseur", "Désignation"])

    feuille_sans_catalogue = _mettre_en_colonnes(
        d[sans_catalogue], motif[sans_catalogue])
    feuille_sans_catalogue = feuille_sans_catalogue.sort_values(
        ["Fournisseur", "Désignation"])

    return (feuille_ecarte, feuille_fournisseur_arrete, feuille_arret,
            feuille_sans_catalogue, d)


def main() -> None:
    ecarte, fournisseur_arrete, arret, sans_catalogue, d = construire()

    chemin = mise_en_forme.chemin_ecriture(SORTIE)
    with pd.ExcelWriter(chemin, engine="openpyxl") as w:
        ecarte.to_excel(w, sheet_name="Écarté", index=False, startrow=3)
        sans_catalogue.to_excel(w, sheet_name="Sans catalogue", index=False,
                                startrow=3)
        fournisseur_arrete.to_excel(w, sheet_name="Fournisseur arrêté",
                                    index=False, startrow=3)
        arret.to_excel(w, sheet_name="Arrêt fabricant", index=False,
                       startrow=3)

    reste = int((ecarte["Statut PA"].str.startswith("3")).sum())
    mise_en_forme.formater(chemin, {
        "Écarté": {
            "ligne_entete": 4, "figer_colonne": 4,
            "titre": "Écarté — lits, fauteuils, releveurs, VPH",
            "sous_titre": (
                f"{len(ecarte)} articles, dont {reste} encore sans prix | "
                f"mis de côté, PAS abandonnés : le produit se vend toujours. "
                f"« Motif » dit pourquoi chacun est là. Arrêté au "
                f"{date.today():%d/%m/%Y}")},
        "Sans catalogue": {
            "ligne_entete": 4, "figer_colonne": 4,
            "titre": "Sans catalogue — aucun tarif reçu du fournisseur",
            "sous_titre": (
                f"{len(sans_catalogue)} articles | ni choix ni arrêt : il "
                f"manque la matière première. Le produit se vend, le "
                f"fournisseur existe, on n'a rien à lire. Le seul geste "
                f"qui les débloque est de réclamer le tarif. Arrêté au "
                f"{date.today():%d/%m/%Y}")},
        "Fournisseur arrêté": {
            "ligne_entete": 4, "figer_colonne": 4,
            "titre": "Fournisseur arrêté — plus de commande chez eux",
            "sous_titre": (
                f"{len(fournisseur_arrete)} articles | constat des appros, "
                f"daté — voir fournisseurs_arretes.py. Le produit n'a rien, "
                f"c'est la relation commerciale qui s'est arrêtée. Arrêté "
                f"au {date.today():%d/%m/%Y}")},
        "Arrêt fabricant": {
            "ligne_entete": 4, "figer_colonne": 4,
            "titre": "Arrêt fabricant — introuvable au catalogue",
            "sous_titre": (
                f"{len(arret)} articles | uniquement ce que VOUS avez "
                f"annoté : le WMS ne signale aucun arrêt (aucun Type 79 au "
                f"périmètre), rien ici n'est déduit. Arrêté au "
                f"{date.today():%d/%m/%Y}")},
    })

    print(f"  {len(ecarte):>5}  Écarté "
          f"({reste} sans prix, {len(ecarte) - reste} déjà valorisés)")
    for m, n in ecarte["Motif"].str.split(" :").str[0].value_counts().items():
        print(f"         {n:>5}  {m}")
    print(f"  {len(sans_catalogue):>5}  Sans catalogue "
          f"({sans_catalogue['Fournisseur'].nunique()} fournisseurs)")
    for f, n in sans_catalogue["Fournisseur"].replace(
            "", "(aucun fournisseur)").value_counts().head(10).items():
        print(f"         {n:>5}  {f}")
    print(f"  {len(fournisseur_arrete):>5}  Fournisseur arrêté")
    for f, n in fournisseur_arrete["Fournisseur"].value_counts().items():
        print(f"         {n:>5}  {f}")
    print(f"  {len(arret):>5}  Arrêt fabricant")

    orphelins = d[d["Origine du fournisseur"] == "introuvable"]
    repris = int((d["Origine du fournisseur"] == "fabricant (extract WMS)").sum())
    print(f"\n  {repris} fournisseurs repris du fabricant (extract WMS)")
    if len(orphelins):
        print(f"  {len(orphelins)} sans aucun fournisseur NI fabricant :")
        for _, r in orphelins.iterrows():
            print(f"      {r['Code article']:<8} {r['Désignation']}")

    print(f"\n-> {chemin}")


if __name__ == "__main__":
    main()
