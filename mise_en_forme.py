# -*- coding: utf-8 -*-
"""
Mise en forme des classeurs Excel produits.

Le fichier doit etre exploitable des l'ouverture : en-tetes lisibles, volets
figes, filtres prets, largeurs adaptees au contenu reel et formats de nombre
corrects. Les anomalies sont mises en evidence par la couleur pour eviter
d'avoir a les chercher.

Point critique : les identifiants (EAN, references, codes article) sont
forces en format Texte. Sans cela Excel perd les zeros de tete et bascule
les EAN en notation scientifique des la premiere ouverture.
"""

from __future__ import annotations

import re
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# --- palette ---------------------------------------------------------------
BLEU_ENTETE = "2E5B8A"
BLEU_ALTERNE = "F2F6FC"
BLEU_PALIER = "DDEBF7"
GRIS_BORDURE = "D9D9D9"
ROUGE_CLAIR = "FFC7CE"     # prix incoherent
ROUGE_VIF = "FF5B5B"       # prix nul ou negatif : bloquant
ORANGE_CLAIR = "FFE0B2"    # EAN manquant / proposition a arbitrer
VERT_CLAIR = "C6EFCE"      # proposition recommandee

POLICE = "Arial"

# --- formats de nombre -----------------------------------------------------
FORMAT_EURO = '# ##0.00\\ "â‚¬"'
FORMAT_POURCENT = "0.0 %"
FORMAT_ENTIER = "0"
FORMAT_TEXTE = "@"

# Colonnes forcees en texte, quel que soit l'onglet
COLONNES_TEXTE = {
    "fournisseur", "ref_fournisseur", "designation", "ean", "paliers",
    "code_lppr", "dispositif_medical", "origine", "fichier_source",
    "code_article", "code_article_wms", "code_declinaison",
    "reference_interne", "reference_wms", "ref_article_fournisseur",
    "reference_fabricant",
    "designation_wms", "dans_wms", "methode_rapprochement", "ean_valide",
    "type_anomalie", "detail", "diagnostic", "indicateur",
    "statut", "motif", "raison", "palier", "etablissements",
    # liste interne "Ref catalogue pro"
    "reference_distributeur", "designation_distributeur", "fournisseur_devine",
    "classe_dm", "source", "commentaire_quantite", "fichier_refs",
    # demande de codes EAN adressee au fournisseur
    "Produit", "Tailles concernÃ©es", "RÃ©f. Fournisseur B connues",
    "CODE EAN (Ã  complÃ©ter)", "references", "codes_article",
}

COLONNES_EURO = {
    "prix_achat_unitaire_ht", "prix_colis_ht", "tarif_public_ttc",
    "prix_recalcule", "ecart_prix", "palier2_prix_ht", "palier3_prix_ht",
    "eco_part_ht", "montant_lppr",
    # propositions d'achat
    "prix_unitaire_normal", "prix_unitaire_propose", "cout_total_normal",
    "cout_total_propose", "economie_unitaire_eur", "economie_totale_eur",
    "surcout_tresorerie", "prix_unitaire_palier",
}

COLONNES_POURCENT = {
    "tva_taux", "remise_taux", "economie_palier2_pct", "economie_palier3_pct",
    "economie_pct",
}

COLONNES_ENTIER = {
    "conditionnement", "palier2_qte", "palier3_qte", "nb_articles_wms",
    "valeur", "longueur_ean",
    # propositions d'achat
    "qte_normale", "qte_proposee", "stock_actuel", "stock_maxi", "qte_palier",
    # Top 200
    "rang_top200", "ventes_annee",
}

# Colonnes a deux decimales (consommation, couverture)
COLONNES_DECIMAL = {"conso_moyenne_mensuelle", "couverture_mois"}
FORMAT_DECIMAL = "0.00"

# --- largeurs --------------------------------------------------------------
LARGEURS = {
    "ref_fournisseur": 14,
    "reference_wms": 14,
    "code_article_wms": 16,
    "ref_article_fournisseur": 16,
    "designation": 45,
    "designation_wms": 45,
    "ean": 16,
    "type_anomalie": 28,
    "detail": 42,
    "motif": 95,
    "raison": 46,
    # demande de codes EAN adressee au fournisseur
    "Produit": 52,
    "Tailles concernÃ©es": 60,
    "Nb de tailles": 12,
    "Nb dÃ©clinaisons": 14,
    "RÃ©f. Fournisseur B connues": 34,
    "CODE EAN (Ã  complÃ©ter)": 24,
    "statut": 14,
    "conso_moyenne_mensuelle": 13,
    "couverture_mois": 12,
    "designation_distributeur": 45,
    "reference_fabricant": 18,
    "reference_distributeur": 14,
    "fournisseur_devine": 16,
    "commentaire_quantite": 52,
    "source": 24,
    "indicateur": 42,
    "paliers": 30,
    "fichier_source": 30,
    "diagnostic": 32,
}
LARGEUR_PRIX = 16
LARGEUR_QTE = 12
LARGEUR_TAUX = 10
LARGEUR_DEFAUT = 18

# --- objets de style partages (un seul exemplaire, pour la memoire) --------
_bordure = Side(style="thin", color=GRIS_BORDURE)
BORDURE = Border(left=_bordure, right=_bordure, top=_bordure, bottom=_bordure)

FONT_ENTETE = Font(name=POLICE, size=10, bold=True, color="FFFFFF")
FONT_DONNEE = Font(name=POLICE, size=10)
FONT_TITRE = Font(name=POLICE, size=14, bold=True)
FONT_SOUS_TITRE = Font(name=POLICE, size=10, italic=True)

FILL_ENTETE = PatternFill("solid", start_color=BLEU_ENTETE)
FILL_ALTERNE = PatternFill("solid", start_color=BLEU_ALTERNE)
FILL_PALIER = PatternFill("solid", start_color=BLEU_PALIER)
FILL_INCOHERENT = PatternFill("solid", start_color=ROUGE_CLAIR)
FILL_BLOQUANT = PatternFill("solid", start_color=ROUGE_VIF)
FILL_EAN_VIDE = PatternFill("solid", start_color=ORANGE_CLAIR)

ALIGN_ENTETE = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_GAUCHE = Alignment(horizontal="left", vertical="center", wrap_text=False)
ALIGN_CENTRE = Alignment(horizontal="center", vertical="center", wrap_text=False)
ALIGN_DROITE = Alignment(horizontal="right", vertical="center", wrap_text=False)


def _largeur(entete: str) -> float:
    """Largeur adaptee au contenu attendu de la colonne."""
    if entete in LARGEURS:
        return LARGEURS[entete]
    if entete in COLONNES_EURO:
        return LARGEUR_PRIX
    if entete in COLONNES_POURCENT:
        return LARGEUR_TAUX
    if entete in COLONNES_ENTIER:
        return LARGEUR_QTE
    return max(LARGEUR_DEFAUT, min(len(str(entete)) + 3, 30))


def _format_et_alignement(entete: str) -> tuple[str | None, Alignment]:
    """Format de nombre et alignement associes a une colonne."""
    if entete in COLONNES_TEXTE:
        return FORMAT_TEXTE, ALIGN_GAUCHE
    if entete in COLONNES_EURO:
        return FORMAT_EURO, ALIGN_DROITE
    if entete in COLONNES_POURCENT:
        return FORMAT_POURCENT, ALIGN_DROITE
    if entete in COLONNES_ENTIER:
        return FORMAT_ENTIER, ALIGN_CENTRE
    if entete in COLONNES_DECIMAL:
        return FORMAT_DECIMAL, ALIGN_CENTRE
    return None, ALIGN_GAUCHE


def formater_feuille(
    feuille,
    ligne_entete: int = 1,
    figer_colonne: int = 1,
    titre: str | None = None,
    sous_titre: str | None = None,
) -> None:
    """Applique la mise en forme complete a une feuille.

    `ligne_entete` permet de reserver des lignes de synthese au-dessus du
    tableau. `figer_colonne` est l'indice (1-based) de la premiere colonne
    qui doit encore defiler : 3 garde reference et designation visibles.
    """
    if feuille.max_row < ligne_entete:
        return

    entetes = [cellule.value for cellule in feuille[ligne_entete]]
    nb_colonnes = len(entetes)
    premiere_donnee = ligne_entete + 1

    # --- lignes de synthese au-dessus du tableau --------------------------
    if titre:
        cellule = feuille.cell(row=1, column=1, value=titre)
        cellule.font = FONT_TITRE
        cellule.alignment = ALIGN_GAUCHE
    if sous_titre:
        cellule = feuille.cell(row=2, column=1, value=sous_titre)
        cellule.font = FONT_SOUS_TITRE
        cellule.alignment = ALIGN_GAUCHE

    # --- en-tete ----------------------------------------------------------
    for index in range(1, nb_colonnes + 1):
        cellule = feuille.cell(row=ligne_entete, column=index)
        cellule.font = FONT_ENTETE
        cellule.fill = FILL_ENTETE
        cellule.alignment = ALIGN_ENTETE
        cellule.border = BORDURE
        feuille.column_dimensions[get_column_letter(index)].width = _largeur(
            cellule.value
        )
    feuille.row_dimensions[ligne_entete].height = 30

    # --- corps ------------------------------------------------------------
    formats = [_format_et_alignement(entete) for entete in entetes]

    for numero in range(premiere_donnee, feuille.max_row + 1):
        # Alternance discrete une ligne sur deux
        alterne = (numero - premiere_donnee) % 2 == 1
        for index in range(1, nb_colonnes + 1):
            cellule = feuille.cell(row=numero, column=index)
            format_nombre, alignement = formats[index - 1]
            cellule.font = FONT_DONNEE
            cellule.alignment = alignement
            cellule.border = BORDURE
            if format_nombre:
                cellule.number_format = format_nombre
            if alterne:
                cellule.fill = FILL_ALTERNE

    # --- volets et filtre -------------------------------------------------
    feuille.freeze_panes = feuille.cell(
        row=premiere_donnee, column=figer_colonne
    ).coordinate
    debut = f"A{ligne_entete}"
    fin = f"{get_column_letter(nb_colonnes)}{feuille.max_row}"
    feuille.auto_filter.ref = f"{debut}:{fin}"


def surligner_anomalies(feuille, ligne_entete: int = 1) -> None:
    """Met en evidence les lignes et cellules problematiques.

    Ne s'applique qu'aux colonnes reellement presentes dans la feuille :
    chaque onglet ne porte pas les memes informations.
    """
    entetes = [cellule.value for cellule in feuille[ligne_entete]]
    index = {nom: position for position, nom in enumerate(entetes, start=1)}
    nb_colonnes = len(entetes)
    premiere_donnee = ligne_entete + 1

    colonne_coherent = index.get("prix_coherent")
    colonne_prix = index.get("prix_achat_unitaire_ht")
    colonne_ean = index.get("ean")
    # Sur l'onglet Anomalies, c'est le type d'anomalie qui donne la gravite :
    # une ligne "prix nul" y figure justement parce que la cellule est vide.
    colonne_type = index.get("type_anomalie")
    # Onglet des propositions d'achat : le statut colore la ligne entiere
    colonne_statut = index.get("statut")
    colonnes_palier = [
        index[nom]
        for nom in ("palier2_qte", "palier2_prix_ht", "economie_palier2_pct",
                    "palier3_qte", "palier3_prix_ht", "economie_palier3_pct")
        if nom in index
    ]

    for numero in range(premiere_donnee, feuille.max_row + 1):
        # Propositions d'achat : vert si recommande, orange si a arbitrer
        if colonne_statut:
            statut = feuille.cell(row=numero, column=colonne_statut).value
            remplissage = {
                "RECOMMANDE": PatternFill("solid", start_color=VERT_CLAIR),
                "A ARBITRER": PatternFill("solid", start_color=ORANGE_CLAIR),
            }.get(statut)
            if remplissage:
                for colonne in range(1, nb_colonnes + 1):
                    feuille.cell(row=numero, column=colonne).fill = remplissage
            continue

        # Prix nul, negatif OU absent : bloquant, toute la ligne en rouge vif.
        # Un prix manquant empeche de commander tout autant qu'un prix a zero.
        bloquant = False
        if colonne_prix:
            valeur = feuille.cell(row=numero, column=colonne_prix).value
            bloquant = valeur is None or (
                isinstance(valeur, (int, float)) and valeur <= 0
            )

        # Prix incoherent : toute la ligne en rouge clair
        incoherent = False
        if colonne_coherent:
            incoherent = feuille.cell(row=numero, column=colonne_coherent).value is False

        # L'onglet Anomalies porte le verdict dans sa premiere colonne
        if colonne_type:
            type_anomalie = feuille.cell(row=numero, column=colonne_type).value
            bloquant = type_anomalie == "Prix d'achat nul ou negatif"
            incoherent = type_anomalie == "Prix incoherent"

        if bloquant or incoherent:
            remplissage = FILL_BLOQUANT if bloquant else FILL_INCOHERENT
            for colonne in range(1, nb_colonnes + 1):
                feuille.cell(row=numero, column=colonne).fill = remplissage
            continue  # la couleur de ligne prime sur les mises en evidence locales

        # Palier disponible : colonnes de palier en bleu clair
        if colonnes_palier:
            prix_palier = index.get("palier2_prix_ht") or index.get("palier3_prix_ht")
            if prix_palier and feuille.cell(row=numero, column=prix_palier).value:
                for colonne in colonnes_palier:
                    feuille.cell(row=numero, column=colonne).fill = FILL_PALIER

        # EAN manquant : cellule EAN en orange
        if colonne_ean and not feuille.cell(row=numero, column=colonne_ean).value:
            feuille.cell(row=numero, column=colonne_ean).fill = FILL_EAN_VIDE


def derniere_version(dossier, nom: str) -> Path | None:
    """La version de reference d'un fichier de sortie.

    Deux pieges se repondent, et il faut les eviter tous les deux.

    Se fier a la date de modification ne marche pas : ouvrir un vieux
    fichier dans Excel suffit a le rajeunir, et il passerait pour le plus
    recent.

    Mais se fier au seul nom canonique ne marche pas non plus : quand ce
    fichier est ouvert dans Excel, `chemin_ecriture` ecrit a cote sous un
    nom horodate, et c'est cette copie qui porte les donnees fraiches.

    On arbitre donc sur l'horodatage inscrit DANS LE NOM des copies, qui
    lui ne ment pas : si la plus recente est posterieure a la derniere
    ecriture du canonique, c'est elle qui fait foi.
    """
    dossier = Path(dossier)
    principal = dossier / nom
    souche, extension = Path(nom).stem, Path(nom).suffix

    from datetime import datetime

    # Une copie n'est prise au serieux que si son nom et sa date de
    # modification concordent. Un ecart important signe une reouverture
    # dans Excel, qui rajeunit le fichier sans rien y changer : c'est le
    # piege qui faisait passer une consolidation de la veille pour la
    # plus fraiche.
    ECART_TOLERE = 3600  # une heure

    candidats = []
    for chemin in dossier.glob(f"{souche} (*){extension}"):
        trouve = re.search(r"\((\d{8})-(\d{4})\)", chemin.name)
        if not trouve:
            continue
        try:
            annonce = datetime.strptime(trouve.group(1) + trouve.group(2),
                                        "%Y%m%d%H%M")
        except ValueError:
            continue
        modifie = chemin.stat().st_mtime
        if abs(modifie - annonce.timestamp()) <= ECART_TOLERE:
            candidats.append((modifie, chemin))

    if principal.exists():
        candidats.append((principal.stat().st_mtime, principal))
    if not candidats:
        return None
    return max(candidats)[1]


def lire(chemin, feuille: str, colonne_temoin: str, limite: int = 8):
    """Lit une de NOS sorties sans se fier a la position de l'en-tete.

    Nos classeurs portent un titre et un sous-titre au-dessus du tableau,
    d'ou les `skiprows=3` qui parsement le depot. Le probleme est qu'une
    simple ouverture dans Excel peut deplacer cette ligne : le 14/09,
    « EAN fiables.xlsx » a ete converti en tableau et une ligne
    « Colonne1, Colonne2… » s'est glissee au-dessus. Un `skiprows` fixe
    lit alors des donnees DECALEES sans lever la moindre erreur — la pire
    espece de defaut, celle qui rend des chiffres au lieu d'une exception.

    On cherche donc l'en-tete par sa colonne temoin, dans les `limite`
    premieres lignes. Un fichier intact est trouve du premier coup ; un
    fichier remanie par Excel l'est aussi. Si la colonne reste
    introuvable, on s'arrete : mieux vaut un arret qu'un tableau faux.
    """
    import pandas as pd

    from extracteurs.base import chemin_lisible

    chemin = Path(chemin)
    for entete in range(limite):
        try:
            df = pd.read_excel(chemin_lisible(chemin), sheet_name=feuille,
                               header=entete, dtype=str)
        except ValueError:
            # Feuille absente : inutile d'essayer les autres lignes.
            raise SystemExit(
                f"Feuille « {feuille} » absente de {chemin.name}")
        df.columns = [str(c).strip() for c in df.columns]
        if colonne_temoin in df.columns:
            if entete != 3:
                print(f"  ! {chemin.name} : en-tête ligne {entete + 1} "
                      f"et non 4 — fichier remanié dans Excel")
            return df
    raise SystemExit(f"Colonne « {colonne_temoin} » introuvable dans "
                     f"{chemin.name} (feuille « {feuille} »). "
                     f"Le fichier a-t-il été remanié ?")


def chemin_ecriture(chemin: Path) -> Path:
    """Chemin reellement utilisable pour ecrire, meme fichier ouvert.

    Excel verrouille les classeurs ouverts. Plutot que d'interrompre un
    traitement de plusieurs minutes, on ecrit a cote sous un nom horodate
    et on le signale : l'utilisateur compare puis remplace.
    """
    from datetime import datetime

    chemin = Path(chemin)
    if not chemin.exists():
        return chemin
    try:
        with open(chemin, "r+b"):
            return chemin
    except PermissionError:
        horodatage = datetime.now().strftime("%Y%m%d-%H%M")
        secours = chemin.with_name(f"{chemin.stem} ({horodatage}){chemin.suffix}")
        print(f"  ! {chemin.name} est ouvert dans Excel")
        print(f"    Ã©criture dans {secours.name}")
        return secours


def formater(
    chemin: Path,
    options: dict[str, dict] | None = None,
) -> None:
    """Formate tous les onglets d'un classeur.

    `options` permet de personnaliser un onglet :
        {"Catalogue": {"ligne_entete": 4, "figer_colonne": 3,
                       "titre": "...", "sous_titre": "..."}}
    """
    options = options or {}
    classeur = openpyxl.load_workbook(chemin)

    for feuille in classeur.worksheets:
        parametres = options.get(feuille.title, {})
        ligne_entete = parametres.get("ligne_entete", 1)
        formater_feuille(
            feuille,
            ligne_entete=ligne_entete,
            figer_colonne=parametres.get("figer_colonne", 1),
            titre=parametres.get("titre"),
            sous_titre=parametres.get("sous_titre"),
        )
        surligner_anomalies(feuille, ligne_entete=ligne_entete)

    classeur.save(chemin)

