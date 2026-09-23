Supplier price list reconciliation pipeline

Python pipeline that reconstructs, from heterogeneous supplier price lists, the EAN barcode and the purchase price of every product reference held by a medical device distributor, then reconciles them against its warehouse management system.

Disclaimer

Representative extract of a larger system. This repository does not run as-is: it contains no confidential data, and some internal modules were deliberately left out. The code is provided to be read, to illustrate the approach and the architecture. All names (client, suppliers, people) and all commercial figures are anonymized or replaced with examples.

The problem

A distributor receives, every year, the price lists of nearly a hundred suppliers, in as many formats: Excel workbooks with columns that never match, native PDFs, scanned PDFs, commercial-terms documents that are not price lists at all. On the other side, its internal product reference base barely knows the EAN codes and carries outdated purchase prices. Reconciling the two by hand is out of reach at this scale, and doing it naively is worse than doing nothing: a reading error does not produce a crash, it produces a wrong price that nobody notices.

Architecture

The pipeline reads in five stages. The names below are the files in this repository.

1. Extraction. One module per supplier, all bound to the same contract: extract(path) -> DataFrame following the common schema.

extracteurs/base.py: the foundation. Output schema, EAN check-digit validation (mod 10, GS1 standard), consistency check between displayed price and recomputed price, tiered-pricing logic. This is the file to start from.
extracteurs/generique.py: fallback extractor, which detects columns by their header rather than by their position.
extracteurs/pdf_base.py, extracteurs/paliers_lignes.py: specialized foundations.
extracteurs/fournisseur_a.py, fournisseur_b_conditions.py, fournisseur_c.py: three examples kept out of the sixty that exist. The second is the most instructive: it reads compound discounts, where it is the cell format and not its value that tells a price apart from a rate.

2. Scope. The working set is frozen once, then read, never recomputed.

perimetre_config.py: the entry rules, written as owned business decisions and not as parameters to tune.
figer_perimetre.py: produces the frozen list and its manifest (source hashes).
perimetre_liste.py: reads it, and refuses to continue if the count does not match.
perimetre.py, normalisation.py: annotation and normalization of references.

3. Cascade matching. Several keys tested in decreasing order of reliability, the key actually used always traced in a column.

pa_completer.py: the core. Four automatic keys, anti-outlier guardrail.
pa_identifiants.py: fifth pass, on damaged identifiers.
pa_correspondances.py: sixth and last, a table of human decisions that only fills empty cells.
pa_libelle.py and par_libelle.py: fuzzy label matching, at a calibrated threshold, read-only.
pa_par_modele.py: matching by model name, when neither reference nor label suffices.
pa_conditionnement.py: converts a pack price to a unit price, never without corroboration.

4. Control and reporting.

qualite.py: building and ranking of anomalies.
pa_reconcilier_responsable.py: confronts our reading with that of a second, independent pipeline that reads the same price lists.
ean_controle.py: four levels of control on the retained codes, read-only.
controle_calculs.py: walks the calculation chain line by line, without writing anything.
pa_fusionner.py, pa_etat_par_article.py, pa_ecartes.py, pa_a_arbitrer.py: non-destructive merge and review views.
mise_en_forme.py: workbook formatting, and writing that does not break a long run when the target file is open elsewhere.

5. EAN chain and external sources.

decoder_gs1.py: decoding of GS1-128 and DataMatrix, distinguishing the selling unit from the case.
ean_global.py: multi-source consolidation, with an explicit priority order.
ean_version.py: freezing of a numbered version, arbitration of ambiguous codes.
eudamed.py, eudamed_par_reference.py, eudamed_fiabilite.py, ean_verifier_libelle.py: querying of the public European medical device registry, and measurement of how much to trust its answers.
ar/base.py, ar/fournisseur_d.py, extraire_ar.ps1: reading of supplier order acknowledgements, which give prices actually invoiced, at a date.

Orchestration. run_pa.py chains the stages, stops at the first blocking failure, and refuses to start if the reference data is inconsistent.

Engineering highlights
Anti-outlier guardrails. Beyond a certain ratio to the reference price, a value is not injected but diverted to a verification column. A discrepancy of an order of magnitude is not a commercial negotiation, it is a reading error, and a reading error does not flag itself.
No silent deletion. The full population enters the pipeline, and every filter becomes a column. A discarded row carries its reason and stays consultable. A filter applied upstream shows up only in the total: you no longer know what was removed, why, or how much.
Measured calibration, not by guesswork. The fuzzy-label threshold was set on a test set built from already-known matches, masking the reference to keep only the label. The measurement showed that the threshold considered at the outset was unreachable on the cases that mattered.
Detection of a silent identifier-conversion bug. A data-processing library converts a column of identifiers into floating-point numbers as soon as one value is missing. The join then drops to zero without raising any error. The corresponding normalization is applied systematically.
Reconciliation against a third-party pipeline. A second pipeline, written by someone else, reads the same price lists. The two readings are confronted automatically. An early version of this control compared, because of an overly strict join key, a tiny fraction of the rows while concluding there was no divergence: an unearned green light, which is the worst possible outcome.
What is not here, and why
The data. Supplier price lists, warehouse-system exports, working workbooks, order acknowledgements: none of it is published. The .gitignore excludes the corresponding formats.
The commercial-terms registry. The central module that maps each supplier to its price-list folder, its negotiated discounts and its reading quirks is not included: it is commercial information. It is imported by a good share of the files present, which explains part of the unresolved imports.
The field-collection app. The barcode-scanning server and interface used in the warehouse (camera or USB scanner input) were moved to a separate repository, their life cycle not being that of the pipeline.
The OCR chain. Some price lists exist only as image PDFs and require OCR. The module that handles it is not part of this extract; the files here handle text-layer PDFs. The known limits of the approach are documented in the code: a misread character does not produce an error but a wrong, silent price, which is why it crosses two engines and double-checks against a plausibility control.
Tech stack

Python 3, pandas and numpy for tabular processing, openpyxl for reading and writing Excel workbooks, pdfplumber for text-layer PDFs, rapidfuzz for fuzzy label matching, requests for querying public APIs. A PowerShell script for extracting attachments from a local mailbox.

Context

Project designed and built independently as part of a procurement and supply-chain assignment, across several thousand references and nearly 90 suppliers. Usable EAN coverage was raised to over 70 % of the tracked scope, and up-to-date purchase-price coverage to over 75 %.

License

Code provided for demonstration purposes. All rights reserved.

Project content
Almadia
Created by you
Add PDFs, documents, or other text to reference in this project.
Content
DOCUMENTATION_TECHNIQUE.md

MD

DOCUMENTATION_TECHNIQUE.md

MD

DOCUMENTATION_TECHNIQUE.md

MD

DOCUMENTATION_TECHNIQUE.md

MD

rapport_pa (1).docx

DOCX

DOCUMENTATION_EAN.md

MD

rapport_ean.docx

DOCX

# Pipeline de réconciliation catalogue fournisseurs — EAN & prix d'achat ## 1. Résumé projet Un distributeur de dispositifs médicaux disposait de près d'une centaine de tarifs fournisseurs hétérogènes (Excel, PDF natif, PDF scanné) et d'un référentiel produit interne où le code-barres (EAN)

PASTED

GROUPE 1 — À garder : le cœur Le socle d'extraction Fichier Pourquoi le montrer extracteurs/base.py Le fichier phare : schéma commun imposé à tous les extracteurs, clé de contrôle EAN (mod-10 GS1), normalisation des codes, tolérance de cohérence prix, validation des paliers dégressifs, lecture d'

PASTED

CAVIARDAGE_RAPPORT.md

196 lines

MD

Les deux fichiers sont créés. Aucun em dash ni en dash dans le README (vérifié caractère par caractère). Aucune autre modification. .gitignore exclut CAVIARDAGE_RAPPORT.md, les artefacts Python, les journaux, tous les formats de données (.xlsx, .xls, .csv, .pdf, .json, .msg, .html et quelques aut

PASTED

# Pipeline de réconciliation de tarifs fournisseurs Pipeline Python qui reconstitue, à partir de tarifs fournisseurs hétérogènes, le code EAN et le prix d'achat de chaque référence d'un distributeur de dispositifs médicaux, puis les rapproche de son système de gestion d'entrepôt. --- > ### Averti

PASTED
