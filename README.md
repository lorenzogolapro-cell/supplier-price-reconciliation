# Supplier price list reconciliation pipeline

A Python pipeline that reconstructs the EAN barcode and the purchase price of every
product reference held by a medical device distributor, starting from heterogeneous
supplier price lists, and matches them against the distributor's warehouse
management system.

---

> ### Please read first
>
> **This is a representative extract of a larger system.** The repository does not
> run as is: it contains no confidential data, and some internal modules were
> deliberately left out. The code is published to be read, to show the approach and
> the architecture. Every name (client, suppliers, people) and every commercial
> figure has been anonymised or replaced with an example value.

---

## The problem

A distributor receives price lists from close to a hundred suppliers every year, in
as many formats: Excel workbooks whose columns never line up, native PDFs, scanned
PDFs, and commercial terms documents that are not price lists at all. On the other
side, its internal product reference system has patchy barcode coverage and carries
outdated purchase prices. Reconciling the two by hand is out of reach at this scale,
and doing it naively is worse than not doing it: a misread value does not crash
anything, it produces a wrong price that nobody notices.

## Architecture

The pipeline reads in five stages. The names below are the actual files in this
repository.

**1. Extraction.** One module per supplier, all bound by the same contract:
`extract(path) -> DataFrame` conforming to the shared schema.

- `extracteurs/base.py`: the foundation. Output schema, EAN check digit validation
  (mod 10, GS1 standard), coherence check between the displayed price and the
  recalculated one, tiered pricing logic. This is the file to start with.
- `extracteurs/generique.py`: fallback extractor, which locates columns by their
  header text rather than by position.
- `extracteurs/pdf_base.py`, `extracteurs/paliers_lignes.py`: specialised
  foundations.
- `extracteurs/fournisseur_a.py`, `fournisseur_b_conditions.py`, `fournisseur_c.py`:
  three examples kept out of the sixty or so that exist. The second is the most
  instructive: it reads compound discounts, where the cell *format* rather than its
  value is what separates a price from a rate.

**2. Scope.** The working population is frozen once, then read, never recomputed.

- `perimetre_config.py`: the inclusion rules, written as settled business decisions
  rather than as parameters to be tuned.
- `figer_perimetre.py`: produces the frozen list and its manifest (source
  fingerprints).
- `perimetre_liste.py`: reads it, and refuses to continue if the count does not
  match.
- `perimetre.py`, `normalisation.py`: annotation and reference normalisation.

**3. Matching cascade.** Several keys tried in decreasing order of reliability, with
the key that succeeded always recorded in its own column.

- `pa_completer.py`: the core. Four automatic keys, plus an anti-aberration guard.
- `pa_identifiants.py`: fifth pass, for damaged identifiers.
- `pa_correspondances.py`: sixth and last, a table of human decisions that only ever
  fills empty cells.
- `pa_libelle.py` and `par_libelle.py`: fuzzy matching on product labels, with a
  calibrated threshold, read only.
- `pa_par_modele.py`: matching on model name, when neither reference nor label is
  enough.
- `pa_conditionnement.py`: converts a pack price into a unit price, never without
  corroboration.

**4. Control and reporting.**

- `qualite.py`: builds and ranks anomalies.
- `pa_reconcilier_responsable.py`: compares our reading against a second,
  independent pipeline that reads the same price lists.
- `ean_controle.py`: four levels of control over the retained barcodes, read only.
- `controle_calculs.py`: walks the pricing chain line by line, writing nothing.
- `pa_fusionner.py`, `pa_etat_par_article.py`, `pa_ecartes.py`, `pa_a_arbitrer.py`:
  non destructive merge and human review views.
- `mise_en_forme.py`: workbook formatting, and a write path that does not break a
  long running job when the target file is open elsewhere.

**5. Barcode chain and external sources.**

- `decoder_gs1.py`: GS1-128 and DataMatrix decoding, distinguishing a sales unit
  from a case.
- `ean_global.py`: multi source consolidation with an explicit priority order.
- `ean_version.py`: freezes a numbered version, arbitrates ambiguous codes.
- `eudamed.py`, `eudamed_par_reference.py`, `eudamed_fiabilite.py`,
  `ean_verifier_libelle.py`: queries against the public European medical device
  registry, and a measured rule for how far its answers can be trusted.
- `ar/base.py`, `ar/fournisseur_d.py`, `extraire_ar.ps1`: parsing of supplier order
  acknowledgements, which give prices actually invoiced, on a given date.

**Orchestration.** `run_pa.py` chains the stages, stops at the first blocking
failure, and refuses to start if the reference data is inconsistent.

## Engineering notes

- **Anti-aberration guards.** Beyond a certain ratio against the reference price, a
  value is not injected but diverted into a review column. An order of magnitude gap
  is not a commercial negotiation, it is a parsing error, and a parsing error does
  not announce itself.

- **No silent deletion.** The full population enters the pipeline, and every filter
  becomes a column. An excluded row carries its reason and stays visible. A filter
  applied upstream only shows up in the total: you no longer know what was removed,
  why, or how much.

- **Thresholds measured, not guessed.** The fuzzy label matching threshold was set
  against a test set built from already known matches, hiding the reference and
  keeping only the label. Measurement showed that the threshold originally
  considered was unreachable on the cases that mattered.

- **A silent identifier coercion bug, found and contained.** A data processing
  library converts a column of identifiers into floating point numbers as soon as
  one value is missing. The join then returns nothing, without raising anything. The
  corresponding normalisation is applied systematically.

- **Reconciliation against a third party pipeline.** A second pipeline, written by
  someone else, reads the same price lists. Both readings are compared
  automatically. An early version of that check compared a tiny fraction of the rows
  because of an overly strict join key, while reporting no divergence at all: an
  unearned green light, which is the worst possible outcome.

## What is not here, and why

- **The data.** Supplier price lists, warehouse system exports, working workbooks,
  order acknowledgements: none of it is published. The `.gitignore` excludes the
  corresponding formats.

- **The commercial terms registry.** The central module mapping each supplier to its
  price list folder, its negotiated discounts and its parsing quirks is not
  included: it is commercial information. It is imported by a good share of the
  files present here, which accounts for part of the unresolved imports.

- **The warehouse scanning app.** The server and interface used to scan barcodes on
  the warehouse floor (camera or USB barcode gun) were moved to a separate
  repository, since their lifecycle is not the pipeline's.

- **The OCR chain.** Some price lists exist only as scanned images and require OCR.
  The module that handles this is not part of this extract; the files here deal with
  PDFs that carry a text layer. The known limits of that approach are documented in
  the code: a misread character does not raise an error, it produces a wrong and
  silent price, which is why two engines are cross checked and the result is tested
  for plausibility.

## Stack

Python 3, `pandas` and `numpy` for tabular processing, `openpyxl` for reading and
writing Excel workbooks, `pdfplumber` for text layer PDFs, `rapidfuzz` for fuzzy
label matching, `requests` for public API queries. A PowerShell script handles
attachment extraction from a local mail client.

Identifiers, function names and output column headers remain in French: they are
contracts shared across files and with the workbooks produced, and renaming them
would have been a cosmetic change with a real risk of silent breakage.

## Context

Designed and built independently as part of a procurement and supply chain
assignment, covering several thousand product references and close to 90 suppliers.
Usable barcode coverage was raised to more than 70 % of the tracked scope, and up to
date purchase price coverage to more than 75 %.

## License

Code provided for demonstration purposes. All rights reserved.
