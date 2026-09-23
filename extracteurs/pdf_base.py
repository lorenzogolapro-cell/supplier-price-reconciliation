# -*- coding: utf-8 -*-
"""
PLACEHOLDER for supplier catalogues delivered as PDF.

Nothing is implemented here until a real PDF has been supplied: the
extraction strategy (tables detected by pdfplumber, splitting by column
positions, or line by line reconstruction) depends entirely on the layout
of the document.

Planned skeleton, once a PDF is available:

    import pdfplumber
    from extracteurs.base import finaliser, nettoyer_ean, nettoyer_nombre, nettoyer_texte

    FOURNISSEUR = "NOM_DU_FOURNISSEUR"

    def extract(path):
        lignes = []
        with pdfplumber.open(path) as pdf:
            for page in pdf.pages:
                for table in page.extract_tables():
                    # 1. locate / skip the page header row
                    # 2. map the columns onto the common schema
                    # 3. ignore rows with neither reference nor label
                    ...
        return finaliser(lignes)

The contract is the same as for the Excel extractors: an extract(path)
function returning a pandas.DataFrame that matches
extracteurs.base.COLONNES. The rest of the chain (consolidation, quality
report) does not change.
"""
