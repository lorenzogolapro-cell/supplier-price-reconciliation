# -*- coding: utf-8 -*-
"""
Formatting of the Excel workbooks produced.

The file has to be usable the moment it opens: readable headers, frozen
panes, filters ready, widths matched to the actual content and correct
number formats. Anomalies are highlighted with colour so that nobody has
to go looking for them.

Critical point: identifiers (EAN codes, references, item codes) are
forced to Text format. Without that, Excel drops leading zeros and flips
EAN codes into scientific notation on the very first opening.
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
ROUGE_CLAIR = "FFC7CE"     # inconsistent price
ROUGE_VIF = "FF5B5B"       # zero or negative price: blocking
ORANGE_CLAIR = "FFE0B2"    # missing EAN / proposal to arbitrate
VERT_CLAIR = "C6EFCE"      # recommended proposal

POLICE = "Arial"

# --- number formats --------------------------------------------------------
FORMAT_EURO = '# ##0.00\\ "€"'
FORMAT_POURCENT = "0.0 %"
FORMAT_ENTIER = "0"
FORMAT_TEXTE = "@"

# Columns forced to text, whatever the tab
COLONNES_TEXTE = {
    "fournisseur", "ref_fournisseur", "designation", "ean", "paliers",
    "code_lppr", "dispositif_medical", "origine", "fichier_source",
    "code_article", "code_article_wms", "code_declinaison",
    "reference_interne", "reference_wms", "ref_article_fournisseur",
    "reference_fabricant",
    "designation_wms", "dans_wms", "methode_rapprochement", "ean_valide",
    "type_anomalie", "detail", "diagnostic", "indicateur",
    "statut", "motif", "raison", "palier", "etablissements",
    # internal list "Ref catalogue pro"
    "reference_distributeur", "designation_distributeur", "fournisseur_devine",
    "classe_dm", "source", "commentaire_quantite", "fichier_refs",
    # request for EAN codes sent to the supplier
    "Produit", "Tailles concernées", "Réf. Fournisseur B connues",
    "CODE EAN (à compléter)", "references", "codes_article",
}

COLONNES_EURO = {
    "prix_achat_unitaire_ht", "prix_colis_ht", "tarif_public_ttc",
    "prix_recalcule", "ecart_prix", "palier2_prix_ht", "palier3_prix_ht",
    "eco_part_ht", "montant_lppr",
    # purchase proposals
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
    # purchase proposals
    "qte_normale", "qte_proposee", "stock_actuel", "stock_maxi", "qte_palier",
    # Top 200
    "rang_top200", "ventes_annee",
}

# Two-decimal columns (consumption, coverage)
COLONNES_DECIMAL = {"conso_moyenne_mensuelle", "couverture_mois"}
FORMAT_DECIMAL = "0.00"

# --- widths ----------------------------------------------------------------
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
    # request for EAN codes sent to the supplier
    "Produit": 52,
    "Tailles concernées": 60,
    "Nb de tailles": 12,
    "Nb déclinaisons": 14,
    "Réf. Fournisseur B connues": 34,
    "CODE EAN (à compléter)": 24,
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

# --- shared style objects (one instance only, to save memory) --------------
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
    """Width matched to the column's expected content."""
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
    """Number format and alignment attached to a column."""
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
    """Applies the full formatting to a sheet.

    `ligne_entete` makes it possible to reserve summary rows above the
    table. `figer_colonne` is the (1-based) index of the first column
    that should still scroll: 3 keeps reference and description visible.
    """
    if feuille.max_row < ligne_entete:
        return

    entetes = [cellule.value for cellule in feuille[ligne_entete]]
    nb_colonnes = len(entetes)
    premiere_donnee = ligne_entete + 1

    # --- summary rows above the table -------------------------------------
    if titre:
        cellule = feuille.cell(row=1, column=1, value=titre)
        cellule.font = FONT_TITRE
        cellule.alignment = ALIGN_GAUCHE
    if sous_titre:
        cellule = feuille.cell(row=2, column=1, value=sous_titre)
        cellule.font = FONT_SOUS_TITRE
        cellule.alignment = ALIGN_GAUCHE

    # --- header -----------------------------------------------------------
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

    # --- body -------------------------------------------------------------
    formats = [_format_et_alignement(entete) for entete in entetes]

    for numero in range(premiere_donnee, feuille.max_row + 1):
        # Discreet banding every other row
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

    # --- panes and filter -------------------------------------------------
    feuille.freeze_panes = feuille.cell(
        row=premiere_donnee, column=figer_colonne
    ).coordinate
    debut = f"A{ligne_entete}"
    fin = f"{get_column_letter(nb_colonnes)}{feuille.max_row}"
    feuille.auto_filter.ref = f"{debut}:{fin}"


def surligner_anomalies(feuille, ligne_entete: int = 1) -> None:
    """Highlights the problematic rows and cells.

    Only applies to the columns actually present in the sheet: not every
    tab carries the same information.
    """
    entetes = [cellule.value for cellule in feuille[ligne_entete]]
    index = {nom: position for position, nom in enumerate(entetes, start=1)}
    nb_colonnes = len(entetes)
    premiere_donnee = ligne_entete + 1

    colonne_coherent = index.get("prix_coherent")
    colonne_prix = index.get("prix_achat_unitaire_ht")
    colonne_ean = index.get("ean")
    # On the Anomalies tab, it is the anomaly type that gives the severity:
    # a "zero price" row is there precisely because the cell is empty.
    colonne_type = index.get("type_anomalie")
    # Purchase proposals tab: the status colours the whole row
    colonne_statut = index.get("statut")
    colonnes_palier = [
        index[nom]
        for nom in ("palier2_qte", "palier2_prix_ht", "economie_palier2_pct",
                    "palier3_qte", "palier3_prix_ht", "economie_palier3_pct")
        if nom in index
    ]

    for numero in range(premiere_donnee, feuille.max_row + 1):
        # Purchase proposals: green if recommended, orange if to arbitrate
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

        # Zero, negative OR missing price: blocking, whole row in bright red.
        # A missing price blocks ordering just as much as a price of zero.
        bloquant = False
        if colonne_prix:
            valeur = feuille.cell(row=numero, column=colonne_prix).value
            bloquant = valeur is None or (
                isinstance(valeur, (int, float)) and valeur <= 0
            )

        # Inconsistent price: whole row in light red
        incoherent = False
        if colonne_coherent:
            incoherent = feuille.cell(row=numero, column=colonne_coherent).value is False

        # The Anomalies tab carries the verdict in its first column
        if colonne_type:
            type_anomalie = feuille.cell(row=numero, column=colonne_type).value
            bloquant = type_anomalie == "Prix d'achat nul ou negatif"
            incoherent = type_anomalie == "Prix incoherent"

        if bloquant or incoherent:
            remplissage = FILL_BLOQUANT if bloquant else FILL_INCOHERENT
            for colonne in range(1, nb_colonnes + 1):
                feuille.cell(row=numero, column=colonne).fill = remplissage
            continue  # the row colour wins over local highlights

        # Price break available: price-break columns in light blue
        if colonnes_palier:
            prix_palier = index.get("palier2_prix_ht") or index.get("palier3_prix_ht")
            if prix_palier and feuille.cell(row=numero, column=prix_palier).value:
                for colonne in colonnes_palier:
                    feuille.cell(row=numero, column=colonne).fill = FILL_PALIER

        # Missing EAN: EAN cell in orange
        if colonne_ean and not feuille.cell(row=numero, column=colonne_ean).value:
            feuille.cell(row=numero, column=colonne_ean).fill = FILL_EAN_VIDE


def derniere_version(dossier, nom: str) -> Path | None:
    """The reference version of an output file.

    Two traps answer each other, and both have to be avoided.

    Trusting the modification date does not work: simply opening an old
    file in Excel is enough to make it look younger, and it would then
    pass for the most recent one.

    But trusting the canonical name alone does not work either: when
    that file is open in Excel, `chemin_ecriture` writes next to it under
    a timestamped name, and it is that copy which carries the fresh data.

    So the decision is made on the timestamp written INTO THE NAME of the
    copies, which does not lie: if the most recent one is later than the
    last write of the canonical file, it is the one that is authoritative.
    """
    dossier = Path(dossier)
    principal = dossier / nom
    souche, extension = Path(nom).stem, Path(nom).suffix

    from datetime import datetime

    # A copy is only taken seriously if its name and its modification
    # date agree. A large gap is the signature of a reopening in Excel,
    # which makes the file look younger without changing anything in it:
    # that is the trap that made yesterday's consolidation pass for the
    # freshest one.
    ECART_TOLERE = 3600  # one hour

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
    """Reads one of OUR outputs without trusting the header position.

    Our workbooks carry a title and a subtitle above the table, hence the
    `skiprows=3` scattered across the repository. The problem is that
    simply opening the file in Excel can move that row: on 14/09,
    "EAN fiables.xlsx" was converted into a table and a
    "Colonne1, Colonne2..." row slipped in above. A fixed `skiprows` then
    reads SHIFTED data without raising the slightest error, the worst
    kind of defect, the one that returns figures instead of an exception.

    So we look for the header by its witness column, within the first
    `limite` rows. An intact file is found on the first try; a file
    reworked by Excel is found too. If the column stays missing, we stop:
    better a halt than a wrong table.
    """
    import pandas as pd

    from extracteurs.base import chemin_lisible

    chemin = Path(chemin)
    for entete in range(limite):
        try:
            df = pd.read_excel(chemin_lisible(chemin), sheet_name=feuille,
                               header=entete, dtype=str)
        except ValueError:
            # Sheet missing: no point trying the other rows.
            raise SystemExit(
                f"Sheet '{feuille}' missing from {chemin.name}")
        df.columns = [str(c).strip() for c in df.columns]
        if colonne_temoin in df.columns:
            if entete != 3:
                print(f"  ! {chemin.name}: header on row {entete + 1} "
                      f"and not 4, file reworked in Excel")
            return df
    raise SystemExit(f"Column '{colonne_temoin}' not found in "
                     f"{chemin.name} (sheet '{feuille}'). "
                     f"Has the file been reworked?")


def chemin_ecriture(chemin: Path) -> Path:
    """A path that can really be written to, even with the file open.

    Excel locks open workbooks. Rather than aborting a job that takes
    several minutes, we write next to it under a timestamped name and say
    so: the user compares, then replaces.
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
        print(f"  ! {chemin.name} is open in Excel")
        print(f"    writing to {secours.name}")
        return secours


def formater(
    chemin: Path,
    options: dict[str, dict] | None = None,
) -> None:
    """Formats every tab of a workbook.

    `options` allows a tab to be customised:
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

