# -*- coding: utf-8 -*-
"""
EMPLACEMENT RESERVE pour les catalogues fournisseurs au format PDF.

Rien n'est implemente ici tant qu'un vrai PDF n'a pas ete fourni : la
strategie d'extraction (tables detectees par pdfplumber, decoupage par
positions de colonnes, ou reconstruction ligne par ligne) depend entierement
de la mise en page du document.

Squelette prevu, une fois un PDF disponible :

    import pdfplumber
    from extracteurs.base import finaliser, nettoyer_ean, nettoyer_nombre, nettoyer_texte

    FOURNISSEUR = "NOM_DU_FOURNISSEUR"

    def extract(path):
        lignes = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables():
                    # 1. reperer / sauter la ligne d'en-tete de la page
                    # 2. mapper les colonnes vers le schema commun
                    # 3. ignorer les lignes sans reference ni designation
                    ...
        return finaliser(lignes)

Le contrat reste identique a celui des extracteurs Excel : une fonction
extract(path) -> pandas.DataFrame respectant extracteurs.base.COLONNES.
Le reste de la chaine (consolidation, rapport qualite) ne change pas.
"""
