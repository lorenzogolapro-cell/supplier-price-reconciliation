# -*- coding: utf-8 -*-
"""
Chantier PA — rejouer un rapprochement de tarif sur le perimetre fige.

    python pa_completer.py FOURNISSEUR_A
    python pa_completer.py FOURNISSEUR_A FOURNISSEUR_E FOURNISSEUR_I

Le responsable des prix a integre 34 tarifs et rapproche chacun d'une
population etablie AVANT le gel du perimetre. Sur le Fournisseur A : il a
traite 576 articles, le perimetre en compte 686, et sa fiche signale 937
« PA hors referencement » — des prix trouves au tarif puis laisses de
cote parce que l'article n'etait pas dans sa liste.

Ce script ne parse rien de neuf. Il reprend l'extracteur existant du
fournisseur et le rapproche des articles du perimetre, pour recuperer ce
gisement deja paye.

Il suit la logique du responsable des prix, colonne pour colonne, afin
que sa feuille puisse etre completee sans retraitement :

    Ancien PA | Nouveau PA | Ecart % | Origine Nouveau PA
    Fichier source PA | Alerte

Le responsable des prix fait foi sur les PA (PERIMETRE.md §6) : ce
fichier PROPOSE. Il ne remplace jamais une valeur qu'il a deja posee —
les articles qu'il a traites sont marques et laisses tels quels.

Sortie : sortie/3-achats/pa_completer_<FOURNISSEUR>.xlsx
    "A completer"   les articles du perimetre sans PA, avec le prix propose
    "Deja traite"   ce que le responsable des prix a deja fait, pour
                    controle
    "Sans tarif"    articles du perimetre absents du tarif fournisseur
"""

from __future__ import annotations

import re
import sys
from functools import lru_cache
from pathlib import Path

import pandas as pd

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import eudamed_fiabilite  # noqa: E402
import mise_en_forme  # noqa: E402
import pa_conditionnement  # noqa: E402
import pa_correspondances  # noqa: E402
import pa_identifiants  # noqa: E402
import pa_relecture  # noqa: E402
import pa_remises  # noqa: E402
import perimetre_config  # noqa: E402
import perimetre_liste  # noqa: E402
import prix_achat_wms  # noqa: E402
from achats_par_article import mots_distinctifs  # noqa: E402
from extracteurs import generique  # noqa: E402
from extracteurs.base import (  # noqa: E402
    chemin_lisible,
    cle_ref_souple,
    cle_ref_stricte,
)
from main import (  # noqa: E402
    CATALOGUES,
    FOURNISSEURS,
    SORTIE_ACHATS,
    SORTIE_CHANTIER,
    dernier_prix_achat,
)

INVENTAIRE = SORTIE_CHANTIER / "inventaire_catalogues.xlsx"
from normalisation import normaliser_reference  # noqa: E402

DONNEES = Path("./data")
PERIMETRE = RACINE / "sortie" / "4-perimetre" / "perimetre_v1_1.xlsx"
import wms_extract  # noqa: E402
EXTRACT = wms_extract.chemin()
VERSION = getattr(perimetre_config, "PERIMETRE_VERSION", "v1.1")

# --- seuils d'alerte, calques sur ceux du responsable des prix -----------
# Un ecart modere est attendu : nos prix datent, le tarif est de l'annee.
# Ce qui doit alerter, c'est l'ecart hors de proportion — presque toujours
# un prix de colis face a un prix unitaire.
ECART_A_VERIFIER = 0.30       # au-dela, on demande une relecture
ECART_SUSPECT = 1.00          # au-dela, le prix n'est probablement pas comparable
# Rapports qui trahissent un conditionnement plutot qu'une hausse.
CONDITIONNEMENTS = (2, 3, 4, 5, 6, 10, 12, 20, 24, 25, 50, 100)
TOLERANCE_CONDITIONNEMENT = 0.02

# Au-dela de ce rapport, une hausse de prix n'est plus une explication
# plausible. Chez les fournisseurs de consommable, le WMS stocke souvent
# le prix a l'UNITE (0,0500 EUR la seringue — valeur d'exemple) quand le
# tarif cote au CARTON (5,00 EUR les cent — valeur d'exemple). Le rapport
# exact ne tombe pas sur un
# conditionnement rond — les cartons changent de taille, les remises s'y
# ajoutent — mais la cause est la, et il faut la nommer plutot que de
# se contenter d'un « suspect » qui n'oriente personne.
RAPPORT_CONDITIONNEMENT_PROBABLE = 3.0

# Un ratio IDENTIQUE qui revient sur des dizaines d'articles differents
# n'est pas une hausse de prix : c'est une remise que le tarif ne porte
# pas. Chez le Fournisseur D, 63 articles sortent a 1,2500 — soit
# exactement 1/0,80 : le tarif est brut, notre PA est net d'une remise de
# 20 % (taux et ratio : valeurs d'exemple).
# Une vraie evolution tarifaire ne tombe jamais sur le meme dix-milliemme
# pour autant de references.
# Au-dela de ce rapport avec le PA du WMS, l'ecart ne s'explique plus
# par une remise ou une hausse : c'est une erreur de lecture du tarif.
# Un x109 chez le Fournisseur I venait d'un prix au colis pris pour un prix
# unitaire ; une remise, meme forte, ne depasse pas un facteur 10.
RATIO_ABERRANT = 10.0

REMISE_ARTICLES_MINIMUM = 5      # en deca, c'est une coincidence
REMISE_ECART_MINIMUM = 0.15      # en deca, c'est une hausse ordinaire


def _parfum_conditionnement(rapport) -> int | None:
    """Le conditionnement le plus proche du rapport, s'il y en a un."""
    if rapport is None or pd.isna(rapport) or rapport <= 0:
        return None
    for n in CONDITIONNEMENTS:
        for cible in (n, 1 / n):
            if abs(rapport - cible) <= TOLERANCE_CONDITIONNEMENT * cible:
                return n
    return None


def alerte(ancien, nouveau) -> str:
    """Le message d'alerte, dans la forme employee par le responsable des prix."""
    if nouveau is None or pd.isna(nouveau):
        return ""
    if ancien is None or pd.isna(ancien) or ancien <= 0:
        return "PAS D'ANCIEN PA — première valorisation"

    rapport = nouveau / ancien
    ecart = rapport - 1
    conditionnement = _parfum_conditionnement(rapport)
    if conditionnement and abs(ecart) > ECART_A_VERIFIER:
        return (f"PA fournisseur NON APPLIQUÉ : {nouveau:.4f} € contre "
                f"{ancien:.4f} € actuel (x{rapport:.1f}) — conditionnement "
                f"probablement différent, à vérifier")
    # Rapport trop grand pour etre une hausse : on nomme la cause
    if (rapport >= RAPPORT_CONDITIONNEMENT_PROBABLE
            or rapport <= 1 / RAPPORT_CONDITIONNEMENT_PROBABLE):
        return (f"PA fournisseur NON APPLIQUÉ : {nouveau:.4f} € contre "
                f"{ancien:.4f} € actuel (x{rapport:.1f}) — écart trop fort "
                f"pour une hausse ; prix au colis contre prix unitaire ?")
    if abs(ecart) > ECART_SUSPECT:
        return f"VÉRIF ÉCART {ecart:+.0%} — SUSPECT"
    if abs(ecart) > ECART_A_VERIFIER:
        return f"VÉRIF ÉCART {ecart:+.0%} — INCOHÉRENCE POSSIBLE"
    return ""


# ---------------------------------------------------------------------------

# Le nom WMS et le dossier catalogue n'ont parfois AUCUN mot commun :
# aucune mecanique ne peut les rapprocher, il faut le dire une fois.
ALIAS_CATALOGUE = {
    # A gauche la raison sociale telle que le WMS la porte, a droite le
    # nom du dossier catalogue : les deux n'ont aucun mot commun.
    "FOURNISSEUR AM": "DOSSIER-AM",
    "FOURNISSEUR BX": "DOSSIER-BX",
    # Pour le Fournisseur BK, la raison sociale est celle de la societe
    # et le dossier porte sa MARQUE commerciale : le WMS connait la
    # premiere, le dossier catalogue porte la seconde. Sans cet alias les
    # 27 articles de ce fournisseur passent pour « aucun tarif recu »
    # alors que le tarif 2026 est la.
    "FOURNISSEUR BK": "DOSSIER-BK",
}

# Mots trop repandus dans les raisons sociales du secteur pour identifier
# qui que ce soit. Sans ce filtre, la raison sociale du Fournisseur BV —
# qui contient CONFORT — tombe sur le tarif du Fournisseur CT, qui le
# contient aussi : l'erreur que la fiche de suivi du responsable des prix
# a d'ailleurs commise elle aussi.
MOTS_TROP_COURANTS = {
    "CONFORT", "SANTE", "MEDICAL", "MEDICALE", "FRANCE", "FRANCAISE",
    "SAS", "SARL", "GROUPE", "LABORATOIRE", "LABORATOIRES", "MEDICAUX",
    "INTERNATIONAL", "INTERNATIONALE", "DISTRIBUTION", "TECHNIQUE",
    "SERVICES", "EUROPE", "PHARMA", "SELF", "TOUS",
    # La raison sociale du Fournisseur BV tombait sur le dossier du
    # Fournisseur AX par le seul mot ERGO — ce sont deux societes
    # differentes. ERGO designe une ergonomie, pas une enseigne.
    "ERGO", "CONCEPT",
}


def mots_du_fournisseur(nom) -> set:
    """Mots identifiants d'une raison sociale, mots creux retires."""
    return mots_distinctifs(nom) - MOTS_TROP_COURANTS


def tarif_du_fournisseur(motif_wms: str,
                         noms_reels: tuple[str, ...] = ()) -> pd.DataFrame:
    """Lignes tarifaires du fournisseur, via l'extracteur existant.

    On ne reecrit aucun parseur. Deux familles de fournisseurs :

      - ceux du registre FOURNISSEURS de main.py, qui ont un extracteur
        dedie connaissant leur logique tarifaire ;
      - tous les autres, lus par extracteurs/generique a partir de
        l'inventaire — c'est ainsi que consolider.py prend les
        Fournisseurs E, CS ou CD, absents du registre.
    """
    # Les mots qui identifient ce fournisseur : ceux de son nom WMS
    # reel, plus ceux de son dossier catalogue quand aucun mot n'est
    # commun aux deux (le dossier « DOSSIER-AM » contre la raison sociale
    # du Fournisseur AM).
    attendus = set()
    for nom in noms_reels or (motif_wms,):
        attendus |= mots_du_fournisseur(nom)
    for alias, dossier in ALIAS_CATALOGUE.items():
        if any(alias in str(n).upper() for n in (noms_reels or (motif_wms,))):
            attendus |= mots_du_fournisseur(dossier)
    if not attendus:
        return pd.DataFrame()

    lots = []
    for entree in FOURNISSEURS:
        # Le registre est teste par mots, pas par egalite : son
        # motif_wms est un fragment maison qui ne correspond pas
        # toujours au libelle du WMS.
        if not (mots_du_fournisseur(entree["motif_wms"]) & attendus
                or mots_du_fournisseur(entree["dossier"]) & attendus):
            continue
        dossier = CATALOGUES / entree["dossier"]
        for fichier in sorted(dossier.glob(entree["motif"])):
            if fichier.name.startswith("~$"):
                continue
            try:
                lot = entree["extracteur"].extract(fichier)
            except Exception as err:
                print(f"    ! {fichier.name} : {err}")
                continue
            lot["fichier_tarif"] = fichier.name
            lots.append(lot)
            print(f"    dédié — {fichier.name} : {len(lot)} lignes")
    if lots:
        return pd.concat(lots, ignore_index=True)

    # Repli generique, sur les fichiers que l'inventaire a retenus.
    #
    # Le rapprochement se fait par MOTS DISTINCTIFS, jamais par
    # sous-chaine : « PLAST » se trouve dans la raison sociale du
    # Fournisseur O, et un premier essai a bel et bien applique son tarif
    # au Fournisseur EA, dont la raison sociale contient « PLASTIQUE ».
    # Une sous-chaine ne dit rien de l'identite.
    if not INVENTAIRE.exists():
        print("    ! inventaire absent : lancer inventaire_catalogues.py")
        return pd.DataFrame()
    inventaire = pd.read_excel(chemin_lisible(INVENTAIRE),
                               sheet_name="Retenus", dtype=str)
    excels = inventaire[inventaire["format"].isin(("xlsx", "xlsm", "xls"))]
    for _, ligne in excels.iterrows():
        if not mots_du_fournisseur(ligne["fournisseur"]) & attendus:
            continue
        chemin = Path(ligne["chemin"])
        if not chemin.exists():
            continue
        try:
            lot = generique.extract(chemin, fournisseur=ligne["fournisseur"])
        except Exception as err:
            print(f"    ! {chemin.name} : {type(err).__name__}")
            continue
        if lot.empty:
            continue
        lot["fichier_tarif"] = chemin.name
        lots.append(lot)
        print(f"    générique — {chemin.name} : {len(lot)} lignes")
    if not lots:
        return pd.DataFrame()
    return pd.concat(lots, ignore_index=True)


COMMANDES = DONNEES / "lignes_de_commande.xlsx"


@lru_cache(maxsize=1)
def depuis_les_commandes() -> pd.DataFrame:
    """Fournisseur et reference fournisseur, tires des commandes passees.

    387 articles du perimetre n'ont AUCUN fournisseur sur leur fiche
    article, alors que 277 en ont un — et un seul — sur leurs lignes de
    commande. Le fournisseur n'est pas inconnu, il est ailleurs : sans ce
    repli, ces articles ne se rattachent a aucun tarif et disparaissent
    de tout raisonnement par fournisseur.

    On garde la commande la plus RECENTE : c'est elle qui dit chez qui on
    achete aujourd'hui, un fournisseur ayant pu changer en deux ans.
    """
    if not COMMANDES.exists():
        print("    ! historique de commande absent")
        return pd.DataFrame()
    cmd = pd.read_excel(chemin_lisible(COMMANDES), dtype=str,
                        usecols=["Réf. Art.", "Fournisseur", "Réf. Four.",
                                 "Date de Cde"])
    cmd = cmd.dropna(subset=["Réf. Art.", "Fournisseur"])
    cmd["cle_interne"] = cmd["Réf. Art."].map(normaliser_reference)
    cmd["date"] = pd.to_datetime(cmd["Date de Cde"], errors="coerce")
    # "229 - FOURNISSEUR Y" -> "FOURNISSEUR Y"
    cmd["fournisseur_cmd"] = (cmd["Fournisseur"].str.split("-", n=1)
                              .str[-1].str.strip())
    cmd = cmd.sort_values("date").drop_duplicates("cle_interne", keep="last")
    return cmd[["cle_interne", "fournisseur_cmd", "Réf. Four."]].rename(
        columns={"Réf. Four.": "ref_fournisseur_cmd"})


def references_en_collision(dedans: pd.DataFrame) -> set:
    """Références internes portées par plusieurs articles du périmètre.

    La Reference n'est pas unique : 7 d'entre elles portent deux articles,
    et l'une le fait chez DEUX fournisseurs differents — une gourde chez
    le Fournisseur A et un clip d'electrode chez le Fournisseur AP. Tout
    rapprochement
    par cette clef leur attribuerait la meme chose, et se tromperait une
    fois sur deux.

    On ne tranche pas : on les recense pour les ecarter du repli et les
    signaler en colonne.
    """
    utiles = dedans[dedans["cle_interne"].notna()
                    & (dedans["cle_interne"] != "")]
    comptes = utiles["cle_interne"].value_counts()
    return set(comptes[comptes > 1].index)


@lru_cache(maxsize=1)
def referentiel_enrichi() -> pd.DataFrame:
    """Le perimetre fige, complete de sa reference fournisseur.

    Mis en cache : ce chargement lit trois classeurs de plusieurs
    milliers de lignes, et le script est appele fournisseur par
    fournisseur — sans cache il les relirait a chaque passage.
    """
    # Le perimetre vient de la liste figee, jamais d'un recalcul : deux
    # sessions qui recalculent peuvent diverger sans que personne le voie.
    dedans = perimetre_liste.charger().copy()
    print(f"    {perimetre_liste.entete()}")
    dedans = dedans.rename(columns={"Fournisseur": "Nom fournisseur",
                                    "Désignation": "Libellé déclinaison ^(1)"})
    dedans["cle_interne"] = dedans["Réf. interne"].map(normaliser_reference)

    # La Reference n'est pas unique : on marque les cas plutot que de
    # les laisser fausser un rapprochement en silence.
    collisions = references_en_collision(dedans)
    dedans["collision_reference"] = dedans["cle_interne"].isin(collisions)
    sans_reference = (dedans["cle_interne"].isna()
                      | (dedans["cle_interne"] == ""))
    dedans["sans_reference"] = sans_reference
    if collisions or sans_reference.any():
        print(f"    clé Référence : {len(collisions)} références portées par "
              f"plusieurs articles, {int(sans_reference.sum())} articles sans "
              f"référence — exclus du repli, signalés en colonne")

    # La reference fournisseur vient de l'extrait article : c'est elle
    # qui fait le pont vers le tarif. Le WMS porte AUSSI un jeu fabricant
    # (nom et reference), bien plus complet : 384 des 387 articles sans
    # fournisseur ont un fabricant, et 1 846 des articles sans prix ont
    # une reference fabricant. C'est une seconde porte vers le tarif.
    art = pd.read_excel(chemin_lisible(EXTRACT), dtype=str,
                        usecols=["Code article", "Ref. art. four.",
                                 "Nom fabriquant", "Référence fabricant"])
    pont = art.dropna(subset=["Code article"]).drop_duplicates("Code article")
    codes = pont["Code article"].str.strip()
    cle = dedans["Code article"].str.strip()
    dedans["ref_fournisseur_wms"] = cle.map(
        dict(zip(codes, pont["Ref. art. four."])))
    dedans["ref_fabricant"] = cle.map(
        dict(zip(codes, pont["Référence fabricant"])))
    dedans["nom_fabriquant"] = cle.map(
        dict(zip(codes, pont["Nom fabriquant"])))

    # Repli sur l'historique de commande, pour les fiches incompletes
    cmd = depuis_les_commandes()
    if not cmd.empty:
        dedans = dedans.merge(cmd, on="cle_interne", how="left")
        vide = dedans["Nom fournisseur"].fillna("").str.strip() == ""
        dedans["origine_fournisseur"] = "fiche article"
        # Une reference en collision designe deux articles : le repli
        # leur donnerait le meme fournisseur, faux pour l'un des deux.
        exploitable = ~dedans["collision_reference"] & ~dedans["sans_reference"]
        recuperes = vide & exploitable & dedans["fournisseur_cmd"].notna()
        dedans.loc[recuperes, "Nom fournisseur"] = \
            dedans.loc[recuperes, "fournisseur_cmd"]
        dedans.loc[recuperes, "origine_fournisseur"] = "historique commande"
        # Idem pour la reference, souvent absente elle aussi
        sans_ref = dedans["ref_fournisseur_wms"].fillna("").str.strip() == ""
        repris = sans_ref & exploitable & dedans["ref_fournisseur_cmd"].notna()
        dedans.loc[repris, "ref_fournisseur_wms"] = \
            dedans.loc[repris, "ref_fournisseur_cmd"]
        if recuperes.any() or repris.any():
            print(f"    repli commandes : {int(recuperes.sum())} fournisseurs "
                  f"et {int(repris.sum())} références récupérés")
    else:
        dedans["origine_fournisseur"] = "fiche article"

    # Dernier repli : le fabricant. Il ne DIT PAS chez qui on achete —
    # on peut prendre du Fournisseur B chez un distributeur — mais sans lui
    # ces articles ne sont rattaches a aucun tarif du tout. L'origine
    # reste en colonne pour que la nuance se voie a la relecture.
    vide = dedans["Nom fournisseur"].fillna("").str.strip() == ""
    depuis_fabricant = vide & dedans["nom_fabriquant"].fillna(
        "").str.strip().ne("")
    dedans.loc[depuis_fabricant, "Nom fournisseur"] = \
        dedans.loc[depuis_fabricant, "nom_fabriquant"]
    dedans.loc[depuis_fabricant, "origine_fournisseur"] = \
        "fabricant (fiche article)"
    if depuis_fabricant.any():
        print(f"    repli fabricant : {int(depuis_fabricant.sum())} "
              f"fournisseurs récupérés — attribution à confirmer")
    return dedans


def articles_du_perimetre(motif_wms: str) -> pd.DataFrame:
    """Articles du perimetre fige appartenant a ce fournisseur."""
    dedans = referentiel_enrichi()
    sien = dedans[dedans["Nom fournisseur"].fillna("").str.upper()
                  .str.contains(motif_wms.upper(), regex=False)].copy()
    sien["cle_stricte"] = sien["ref_fournisseur_wms"].map(cle_ref_stricte)
    sien["cle_souple"] = sien["ref_fournisseur_wms"].map(cle_ref_souple)
    sien["cle_fabricant"] = sien["ref_fabricant"].map(cle_ref_stricte)
    return sien


@lru_cache(maxsize=1)
def ean_par_article() -> pd.DataFrame:
    """Code article -> ses codes EAN sûrs, issus du chantier EAN.

    Quatrieme porte vers le tarif, et la plus universelle : un EAN ne
    depend ni de la nomenclature du fournisseur ni de la notre. Beaucoup
    de tarifs le publient alors que leur reference a change.

    On ne retient que les EAN declares SÛRS. Un EAN « probable » qui
    tombe sur la mauvaise ligne du tarif ne produit pas une erreur : il
    produit un prix, faux, et personne ne le verra.

    Deux corrections du 15/09, sur la meme fonction :

    Le fichier. `sorted(glob("EAN distributeur*.xlsx"))[-1]` retenait
    « EAN distributeur.xlsx » — le fichier de TRAVAIL, celui qui change a
    chaque collecte et que la convention interdit de citer. L'ordre
    alphabetique place « . » apres « - », et de toute facon il aurait
    mis « V9 » apres « V16 ». On trie sur le numero.

    Le filtre. « Fiabilité == sûr » laissait entrer 1 321 codes
    EUDAMED obtenus sur une reference trop courte pour designer quoi que
    ce soit. Confrontes au libelle chez EUDAMED meme, 3 sur 370 se sont
    confirmes. Ils ne rapprochaient que 33 prix, dont 5 par un lien
    douteux — la 4e cle sert peu — mais un prix faux et muet est
    exactement ce que cette fonction dit vouloir eviter.
    """
    dossier = RACINE / "sortie" / "2-chantier-ean"
    fichiers = sorted(
        dossier.glob("EAN distributeur - V*.xlsx"),
        key=lambda p: int(re.search(r"- V(\d+)", p.name).group(1)),
    )
    if not fichiers:
        return pd.DataFrame(columns=["Code article", "cle_ean"])
    table = pd.read_excel(chemin_lisible(fichiers[-1]),
                          sheet_name="Codes EAN", skiprows=3, dtype=str)
    table = table[table["Fiabilité"].fillna("").str.strip() == "sûr"]
    table = table[~eudamed_fiabilite.douteux(table)]
    table = table.dropna(subset=["Code article", "Code EAN"])
    liens = pd.DataFrame({
        "Code article": table["Code article"].str.strip(),
        "cle_ean": table["Code EAN"].str.strip(),
    }).drop_duplicates()
    print(f"    référentiel EAN : {len(liens)} liens sûrs "
          f"sur {liens['Code article'].nunique()} articles")
    return liens


@lru_cache(maxsize=1)
def pa_actuels() -> pd.DataFrame:
    """Le Dernier PA du WMS, par code article resolu."""
    chemin = dernier_prix_achat()
    if chemin is None:
        return pd.DataFrame(columns=["code_article", "pa_wms"])
    pa = prix_achat_wms.charger(chemin)
    pa = pa.dropna(subset=["code_article"])
    return pa.drop_duplicates("code_article")[["code_article", "pa_wms"]]


@lru_cache(maxsize=1)
def deja_traites_par_le_responsable() -> frozenset:
    """Clefs internes que le responsable des prix a deja dotees d'un
    Nouveau PA."""
    chemin = DONNEES / "pa_valides.xlsx"
    if not chemin.exists():
        return frozenset()
    valides = pd.read_excel(chemin_lisible(chemin), sheet_name="Feuil1",
                            dtype=str).dropna(subset=["Code article"])
    valides["pa"] = pd.to_numeric(valides["Nouveau PA"], errors="coerce")
    valides = valides[valides["pa"].fillna(0) > 0]

    art = pd.read_excel(chemin_lisible(EXTRACT), dtype=str,
                        usecols=["Code article", "Référence"])
    pont = art.dropna(subset=["Code article"]).drop_duplicates("Code article")
    index = dict(zip(pont["Code article"].str.strip(),
                     pont["Référence"].map(normaliser_reference)))
    return frozenset(valides["Code article"].str.strip().map(index).dropna())


COLONNES_SORTIE = [
    "PERIMETRE_VERSION", "Réf. interne", "Code article",
    "Libellé déclinaison ^(1)", "Nom fournisseur", "origine_fournisseur",
    "Type", "CATEGORIE", "Réf. fournisseur", "Désignation tarif",
    "Ancien PA", "Nouveau PA", "Écart %",
    "Origine Nouveau PA", "Fichier source PA", "Alerte",
    "PA écarté (à vérifier)", "Paliers", "Conditionnement",
    # Un prix ramene a l'unite doit porter la trace de la facon dont il l'a
    # ete : par quel diviseur, sur quelle preuve, et quel etait le prix avant.
    # Sans ces colonnes, une division fausse est indetectable apres coup.
    "diviseur_conditionnement", "source_diviseur", "preuve_diviseur",
    "prix_avant_division", "config_conditionnement", "minimum_de_commande",
    # Et la trace de ce qu'on a REFUSE de diviser : un conditionnement vu
    # dans un libelle puis ecarte parce que le DPA ne le soutenait pas.
    # Sans cette colonne, la ligne ressemble a une ligne qu'aucune regle
    # n'a touchee, et personne ne sait qu'il y a eu arbitrage.
    "division_ecartee",
    # Meme exigence pour la remise absente du tarif : le taux applique, le
    # prix avant remise, et le motif quand elle a ete ecartee.
    # Et quand une FAMILLE echappe au taux general du fournisseur — le
    # MODELE EXEMPLE 1 a 25 % la ou le Fournisseur D est a 20 %, valeurs
    # d'exemple — la colonne le dit avec
    # sa mesure. Une ligne qui ne suit pas la regle generale doit se voir.
    "remise_taux_applique", "prix_avant_remise", "remise_ecartee",
    "remise_famille",
    # La clef n'est pas toujours fiable : on le dit en colonne plutot
    # que de laisser croire a un rapprochement propre.
    "Clé de rapprochement", "preuve_identifiant",
    "collision_reference", "sans_reference",
]


def completer(motif_wms: str,
              articles: pd.DataFrame | None = None) -> dict[str, pd.DataFrame]:
    """Rapproche les articles d'un fournisseur de son tarif.

    `articles` permet de travailler sur une AUTRE population que le
    perimetre fige — le catalogue professionnel, par exemple, dont une
    partie n'a eu ni vente ni reception sur douze mois et n'entre donc
    pas dans le perimetre. Le perimetre reste ce qu'il est : on ne
    l'elargit pas, on lui pose une population a cote, qui porte sa
    propre etiquette.
    """
    print(f"\n{'=' * 74}\n{motif_wms}\n{'=' * 74}")

    # On identifie d'abord les articles, pour savoir de QUEL fournisseur
    # il s'agit vraiment. Chercher le tarif depuis le seul fragment passe
    # en argument, c'est risquer de tomber sur un homonyme.
    if articles is None:
        articles = articles_du_perimetre(motif_wms)
    print(f"  population : {len(articles)} articles")
    if articles.empty:
        print("  aucun article au périmètre — rien à faire ici")
        return {}
    noms = tuple(sorted(articles["Nom fournisseur"].dropna().unique()))
    print(f"  fournisseurs visés : {', '.join(n[:34] for n in noms[:4])}")
    if len(noms) > 1:
        print(f"  ! le fragment « {motif_wms} » désigne {len(noms)} "
              f"fournisseurs distincts — passer un fragment plus précis")

    print("  tarif :")
    tarif = tarif_du_fournisseur(motif_wms, noms)
    if tarif.empty:
        print("  aucun tarif exploitable — rien à faire ici")
        return {}

    # Rapprochement stricte puis souple, comme partout dans le projet
    champs = ["ref_fournisseur", "designation", "prix_achat_unitaire_ht",
              "paliers", "conditionnement", "fichier_tarif", "colonne_prix"]
    champs = [c for c in champs if c in tarif.columns]
    tarif = tarif.copy()
    tarif["cle_stricte"] = tarif["ref_fournisseur"].map(cle_ref_stricte)
    tarif["cle_souple"] = tarif["ref_fournisseur"].map(cle_ref_souple)

    strict = tarif.dropna(subset=["cle_stricte"]).drop_duplicates("cle_stricte")
    fusion = articles.merge(strict[["cle_stricte"] + champs],
                            on="cle_stricte", how="left",
                            suffixes=("", "_tarif"))
    fusion["Clé de rapprochement"] = ""
    fusion.loc[fusion["prix_achat_unitaire_ht"].notna(),
               "Clé de rapprochement"] = "réf. fournisseur"

    # Deux passes de repli. La reference fournisseur assouplie d'abord,
    # puis la reference FABRICANT : sur les articles restes sans prix,
    # elle est renseignee la ou la reference fournisseur manque. Une
    # cle plus faible que la precedente, donc elle est nommee en colonne.
    for cle_article, cle_tarif, etiquette in (
            ("cle_souple", "cle_souple", "réf. fournisseur (souple)"),
            ("cle_fabricant", "cle_stricte", "réf. fabricant")):
        absents = fusion["prix_achat_unitaire_ht"].isna()
        if not absents.any() or cle_article not in articles.columns:
            continue
        source = (tarif.dropna(subset=[cle_tarif])
                  .drop_duplicates(cle_tarif)
                  .rename(columns={cle_tarif: cle_article}))
        appoint = articles[absents.values].merge(
            source[[cle_article] + champs], on=cle_article, how="left",
            suffixes=("", "_tarif"))
        for champ in champs:
            fusion.loc[absents.values, champ] = appoint[champ].values
        neufs = absents & fusion["prix_achat_unitaire_ht"].notna()
        fusion.loc[neufs, "Clé de rapprochement"] = etiquette
        if neufs.any():
            print(f"  {int(neufs.sum())} rapprochés par {etiquette}")

    # Quatrieme passe : l'EAN. Elle ne peut pas se faire comme les trois
    # autres — un article porte souvent PLUSIEURS codes EAN, donc la
    # jointure duplique les lignes. On rapproche a part, on ne garde
    # qu'un tarif par article, puis on reporte par code article.
    absents = fusion["prix_achat_unitaire_ht"].isna()
    liens = ean_par_article()
    if absents.any() and "ean" in tarif.columns and not liens.empty:
        source = tarif.copy()
        source["cle_ean"] = source["ean"].astype(str).str.strip()
        source = source[source["cle_ean"].str.len() >= 8]
        source = source.drop_duplicates("cle_ean")
        if not source.empty:
            candidats = (fusion.loc[absents, ["Code article"]]
                         .merge(liens, on="Code article")
                         .merge(source[["cle_ean"] + champs], on="cle_ean")
                         .drop_duplicates("Code article")
                         .set_index("Code article"))
            vise = absents & fusion["Code article"].isin(candidats.index)
            for champ in champs:
                fusion.loc[vise, champ] = (
                    fusion.loc[vise, "Code article"].map(candidats[champ]))
            fusion.loc[vise, "Clé de rapprochement"] = "code EAN"
            if vise.any():
                print(f"  {int(vise.sum())} rapprochés par code EAN")

    # Cinquieme passe : les identifiants abimes. Elle vient APRES les quatre
    # cles et AVANT tout rapprochement flou, parce qu'un EAN dont la cle de
    # controle tombe juste est une correspondance EXACTE, pas un candidat.
    fusion = pa_identifiants.rapprocher_par_identifiant(fusion, tarif, champs)

    # Sixieme passe : les correspondances qu'un humain a posees a la main,
    # une par une, chacune avec sa preuve.
    fusion = pa_correspondances.appliquer(fusion, tarif, champs, motif_wms)

    # Septieme et derniere passe : les decisions « OUI » de pa_a_relire.xlsx
    # — un lot de candidats de libelle valides ensemble le 17/09, apres
    # filtrage des lots/sets orphelins et des conflations materielles. Meme
    # regle : jamais d'ecrasement d'un prix deja pose.
    fusion = pa_relecture.appliquer(fusion, motif_wms)

    pa = pa_actuels()
    fusion = fusion.merge(pa, left_on=fusion["Code article"].str.strip(),
                          right_on="code_article", how="left")

    # Conditionnement : certains fournisseurs cotent au lot quand le WMS
    # compte a
    # l'unite. Le DPA est joint JUSTE AVANT, parce que la deduction par
    # systematicite en a besoin — c'est le rapport au DPA qui revele un
    # facteur repete. Tout ce qui est divise le dit en colonne.
    fusion = pa_conditionnement.appliquer_au_rapprochement(fusion)

    # Remise absente du tarif : certains fournisseurs publient un prix BRUT
    # et la remise negociee n'apparait nulle part dans le fichier. APRES le
    # conditionnement, jamais avant : on remise un prix deja ramene a
    # l'unite. Le taux ne vient pas d'un document mais des prix que le
    # responsable des prix a arbitres, et il ne s'applique que si le DPA
    # le corrobore.
    fusion = pa_remises.appliquer_au_rapprochement(fusion)

    # Quatre decimales, comme le classeur du responsable des prix : un prix
    # calcule traine sinon ses decimales flottantes (12.3399999999999 —
    # valeur d'exemple), qui ne
    # sont pas un prix mais un artefact.
    fusion["Ancien PA"] = pd.to_numeric(fusion["pa_wms"],
                                        errors="coerce").round(4)
    fusion["Nouveau PA"] = pd.to_numeric(fusion["prix_achat_unitaire_ht"],
                                         errors="coerce").round(4)
    fusion["Écart %"] = (
        (fusion["Nouveau PA"] - fusion["Ancien PA"]) / fusion["Ancien PA"]
    ).round(4)
    fusion["Origine Nouveau PA"] = (
        "Tarif fournisseur " + fusion["fichier_tarif"].fillna("")
    ).where(fusion["Nouveau PA"].notna(), "")
    fusion["Fichier source PA"] = fusion["fichier_tarif"].fillna("")
    fusion["Alerte"] = [alerte(a, n) for a, n
                        in zip(fusion["Ancien PA"], fusion["Nouveau PA"])]

    # Remise systematique : un meme ratio sur beaucoup d'articles dit que
    # le tarif est BRUT et notre PA net. Se detecte a la recurrence, pas
    # a l'ampleur — c'est ce qui la distingue d'une hausse.
    ancien = pd.to_numeric(fusion["Ancien PA"], errors="coerce")
    nouveau = pd.to_numeric(fusion["Nouveau PA"], errors="coerce")
    ratio = (nouveau / ancien).round(2)
    comptes = ratio[ratio.notna() & ((ratio - 1).abs() > REMISE_ECART_MINIMUM)]
    suspects = {r for r, n in comptes.value_counts().items()
                if n >= REMISE_ARTICLES_MINIMUM}
    remise = ratio.isin(suspects)
    if remise.any():
        for r in sorted(suspects):
            n = int((ratio == r).sum())
            fusion.loc[ratio == r, "Alerte"] = (
                f"écart systématique ×{r:.2f} sur {n} articles — le tarif "
                f"est plus cher que le WMS de {1 - 1 / r:.0%} : hausse, ou "
                f"remise que le tarif ne porte pas ?" if r > 1 else
                f"écart systématique ×{r:.2f} sur {n} articles — tarif et "
                f"PA ne semblent pas sur la même base")

    # Le prix du tarif est retenu meme quand l'ecart est systematique :
    # la colonne lue est bien le prix unitaire ou le prix remise, jamais
    # le prix public — c'est la regle d'achat. Un ecart recurrent devient
    # donc une alerte a relire, pas un motif de refus.
    #
    # Reste un garde-fou, car tout ecart n'est pas une remise : un
    # rapport de x109 chez le Fournisseur I venait d'un prix au colis pris pour
    # un prix unitaire. Au-dela de RATIO_ABERRANT, ce n'est plus une
    # question commerciale mais une erreur de lecture, et celle-la ne
    # s'injecte pas.
    aberrant = ratio.notna() & ((ratio > RATIO_ABERRANT)
                                | (ratio < 1 / RATIO_ABERRANT))
    fusion.loc[aberrant, "Alerte"] = (
        "rapport ×" + ratio[aberrant].round(1).astype(str)
        + " avec le PA du WMS — trop grand pour une remise, "
          "prix au colis lu comme un prix unitaire ?")
    fusion["PA écarté (à vérifier)"] = fusion["Nouveau PA"].where(aberrant)
    fusion.loc[aberrant, "Nouveau PA"] = None
    fusion.loc[aberrant, "Écart %"] = None
    ecartes = aberrant
    if remise.any():
        print(f"  {int(remise.sum())} valeurs à écart systématique, "
              f"retenues avec une alerte")
    if aberrant.any():
        print(f"  {int(aberrant.sum())} valeurs écartées "
              f"(rapport aberrant, > ×{RATIO_ABERRANT:.0f})")

    fusion["PERIMETRE_VERSION"] = VERSION
    fusion = fusion.rename(columns={
        "ref_fournisseur": "Réf. fournisseur",
        "designation": "Désignation tarif",
        "paliers": "Paliers",
        "conditionnement": "Conditionnement",
    })

    servies = deja_traites_par_le_responsable()
    fusion["deja_traite"] = fusion["cle_interne"].isin(servies)
    trouve = fusion["Nouveau PA"].notna()
    a_faire = ~fusion["deja_traite"]

    # Un prix ecarte n'est PAS un prix absent du tarif : il y figure,
    # on refuse seulement de l'appliquer tant que le conditionnement
    # n'est pas verifie. Les melanger rendrait les deux feuilles fausses.
    lots = {
        "À compléter": fusion[a_faire & trouve],
        "Déjà traité": fusion[fusion["deja_traite"]],
        "Écartés": fusion[a_faire & ~trouve & ecartes],
        "Sans tarif": fusion[a_faire & ~trouve & ~ecartes],
    }
    # Regle 4 de PERIMETRE.md : le total doit se conserver
    total = sum(len(t) for t in lots.values())
    assert total == len(fusion), f"total non conservé : {total} / {len(fusion)}"

    print(f"  à compléter : {len(lots['À compléter'])}")
    print(f"  déjà traité par le responsable des prix : "
          f"{len(lots['Déjà traité'])}")
    print(f"  absents du tarif : {len(lots['Sans tarif'])}")
    alertes = lots["À compléter"]["Alerte"]
    if len(alertes):
        posees = (alertes != "").sum()
        print(f"  dont {posees} avec une alerte")

    for nom in lots:
        lots[nom] = lots[nom].reindex(
            columns=[c for c in COLONNES_SORTIE if c in fusion.columns])
    return lots


def _nom_fichier(nom: str) -> str:
    """Nom de fichier sûr, et surtout DISTINCT d'un fournisseur à l'autre.

    Prendre les trois premiers mots donnait le meme nom a deux entites
    differentes : « FOURNISSEUR BF FRANCE SITE DE ALPHA » et
    « FOURNISSEUR BF FRANCE SITE DE BETA » produisaient tous deux
    FOURNISSEUR-BF-FRANCE. Le second ecrasait le premier, et le balayage
    croyait ensuite l'avoir traite — 36 articles perdus sans un bruit.

    On garde donc les mots IDENTIFIANTS, ceux qui restent une fois
    retires les mots repandus du secteur : ALPHA et BETA
    survivent, FRANCE et SITE disparaissent.
    """
    mots = [m for m in re.split(r"[^A-Za-z0-9]+", nom.upper()) if m]
    distinctifs = [m for m in mots if m not in MOTS_TROP_COURANTS
                   and len(m) > 2]
    retenus = distinctifs[:4] or mots[:3]
    return "-".join(retenus)[:60] or "FOURNISSEUR"


def fournisseurs_restants() -> list[str]:
    """Fournisseurs du perimetre qu'aucun passage n'a encore couverts.

    Le balayage complet vaut mieux qu'une liste choisie a la main : sur
    les sept fournisseurs qu'on croyait n'avoir qu'un tarif PDF, trois
    avaient en fait un tarif Excel deja lu par le projet. Interroger
    l'extracteur ne coute rien ; deviner coute des articles.
    """
    # Les noms viennent du referentiel ENRICHI, pas du perimetre brut :
    # sinon les articles rattaches par l'historique de commande ou par le
    # fabricant designent un fournisseur que le balayage n'ouvre jamais.
    dedans = referentiel_enrichi()
    noms = (dedans["Nom fournisseur"].dropna().astype(str).str.strip())
    noms = noms[noms != ""]

    # Comparaison par nom de fichier EXACT, pas par mots communs. Le
    # test par intersection considerait « FOURNISSEUR BF ... ALPHA »
    # comme deja traite parce qu'un fichier existait pour « FOURNISSEUR
    # BF ... BETA » : deux entites distinctes, un seul balayage.
    deja = {f.stem.replace("pa_completer_", "").upper()
            for f in SORTIE_ACHATS.glob("pa_completer_*.xlsx")}
    restants = []
    for nom, n in noms.value_counts().items():
        if _nom_fichier(nom).upper() in deja:
            continue
        restants.append(nom)
    print(f"{len(restants)} fournisseurs restants à balayer\n")
    return restants


def main() -> None:
    arguments = sys.argv[1:]
    if arguments and arguments[0] == "--tous":
        cibles = fournisseurs_restants()
    else:
        cibles = arguments or ["FOURNISSEUR_A"]
    for motif in cibles:
        lots = completer(motif)
        if not lots:
            continue
        sortie = mise_en_forme.chemin_ecriture(
            SORTIE_ACHATS / f"pa_completer_{_nom_fichier(motif)}.xlsx")
        with pd.ExcelWriter(sortie, engine="openpyxl") as writeur:
            for nom, table in lots.items():
                table.to_excel(writeur, sheet_name=nom, index=False,
                               startrow=3)
        mise_en_forme.formater(sortie, {
            "À compléter": {
                "ligne_entete": 4, "figer_colonne": 4,
                "titre": f"{motif} — PA à compléter sur le périmètre {VERSION}",
                "sous_titre": "Proposition. Le responsable des prix fait "
                              "foi : on comble, on ne remplace pas."},
            "Déjà traité": {
                "ligne_entete": 4, "figer_colonne": 4,
                "titre": "Déjà doté d'un Nouveau PA par le responsable des prix",
                "sous_titre": "Pour contrôle — ne rien y changer"},
            "Écartés": {
                "ligne_entete": 4, "figer_colonne": 4,
                "titre": "Prix trouvé au tarif mais NON APPLIQUÉ",
                "sous_titre": "Écart trop fort pour une hausse — prix au "
                              "colis contre prix unitaire ? Valeur conservée "
                              "en colonne « PA écarté »"},
            "Sans tarif": {
                "ligne_entete": 4, "figer_colonne": 4,
                "titre": "Au périmètre mais absents du tarif fournisseur",
                "sous_titre": "Ni un bug ni un oubli : le tarif ne les porte pas"},
        })
        print(f"  -> {sortie}")


if __name__ == "__main__":
    main()

