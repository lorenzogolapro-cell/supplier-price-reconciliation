# -*- coding: utf-8 -*-
"""Reconciliation between the price owner's workbook and our purchase-price proposals.

WHY
    Two pipelines read the same price-list files. The price owner's one
    (./data/) covers 33 suppliers whose sheet, header row and exact price-column
    heading were established one by one. Ours sweeps 154 suppliers. Wherever
    both have a price for the same item, the same supplier and the same price
    break, they must say EXACTLY the same thing. Any difference is a reading
    defect on one of the two sides, and we need to know which one before adding
    a single further price.

WHAT THE SCRIPT DOES NOT DO
    It corrects nothing and writes to no working file. It produces a findings
    workbook. The price owner is authoritative: we do not touch their values,
    we document the gap.

READING THE RESULT
    The workbook sortie/3-achats/reconciliation_responsable.xlsx has six tabs:

      Synthese        the counters, to be read first
      Par fournisseur median ratio of ours to the price owner's, supplier by
                      supplier. This is the tab that finds column errors: a
                      stable median ratio over dozens of rows is not thirty
                      disagreements, it is a different column or a different
                      unit.
      Divergences     both have a price and they differ, split by origin
      Seul responsable  the price owner has a price, we do not. Why did our
                      sweep miss it?
      Seul nous       our net contribution, what actually needs validating
      Sans prix       neither side has one, this is the rest of the workstream

USAGE
    python pa_reconcilier_responsable.py [--perimetre chemin.csv]
                                  [--responsable chemin.xlsx]
                                  [--notre chemin.xlsx] [--sortie chemin.xlsx]

    The workbooks can stay open in Excel, they are read from a copy.
"""
import argparse
import os
import re
import shutil
import sys
import tempfile

import numpy as np
import pandas as pd

# Equality tolerance: taken from the convention of verif_ecarts.py.
# Two prices coming from the same cell must match to the cent, not to 2 %.
def _egaux(a, b):
    return (a - b).abs() <= np.maximum(0.005, 0.0005 * b.abs())


# Median ratio per supplier beyond which we call it a wrongly chosen column
# rather than a row-by-row disagreement.
MIN_LIGNES_RATIO = 5
TOLERANCE_RATIO = 0.005


# --------------------------------------------------------------------------- input/output
def chemin_lisible(chemin):
    """Temporary copy: a workbook open in Excel stays readable."""
    if not os.path.exists(chemin):
        raise FileNotFoundError(chemin)
    tmp = os.path.join(tempfile.gettempdir(),
                       "recon_" + os.path.basename(chemin))
    shutil.copy2(chemin, tmp)
    return tmp


def lire(chemin, onglet, temoin="Code article", max_lignes=8):
    """Reads a tab by locating its header row.

    This project's workbooks carry a title block on the first three rows and
    their header on row 4; the price owner's start straight with the header. So
    we look for the first row that contains the witness column rather than
    assuming one convention or the other.
    """
    lisible = chemin_lisible(chemin)
    brut = pd.read_excel(lisible, sheet_name=onglet, header=None,
                         nrows=max_lignes, dtype=str)
    cible = re.sub(r"\s+", " ", temoin).strip().lower()
    for i in range(len(brut)):
        valeurs = {re.sub(r"\s+", " ", str(v)).strip().lower()
                   for v in brut.iloc[i].tolist()}
        if cible in valeurs:
            return pd.read_excel(lisible, sheet_name=onglet, header=i)
    return pd.read_excel(lisible, sheet_name=onglet)


def col(df, *candidats, obligatoire=True, ou=""):
    """Resolves a column name among several possible spellings.

    Fails loudly, listing the columns actually present: that is the only way to
    notice that a tab has been renamed, rather than producing an empty result
    with no error.
    """
    normal = {re.sub(r"\s+", " ", str(c)).strip().lower(): c for c in df.columns}
    for cand in candidats:
        k = re.sub(r"\s+", " ", cand).strip().lower()
        if k in normal:
            return normal[k]
    if not obligatoire:
        return None
    raise KeyError("%s: no column among %s. Present: %s"
                   % (ou or "file", list(candidats), list(df.columns)))


def code(s):
    """Normalises an item code. pandas turns 32818 into "32818.0" as soon as a
    value is missing in the column, and the join then drops to zero without
    raising the slightest error."""
    return (s.astype(str).str.strip()
             .str.replace(r"\.0$", "", regex=True)
             .str.lstrip("/")
             .replace({"nan": np.nan, "None": np.nan, "": np.nan}))


def frs(s):
    """Normalises a supplier label for comparison."""
    return (s.astype(str).str.upper().str.strip()
             .str.replace(r"[^A-Z0-9]+", " ", regex=True)
             .str.replace(r"\s+", " ", regex=True)
             .replace({"NAN": np.nan, "": np.nan}))


def palier(s):
    """Numeric price break. A missing break is NOT a break of 1: the
    distinction is kept, it is what explains part of the gaps."""
    return pd.to_numeric(s, errors="coerce")


# --------------------------------------------------------------------------- loading
def charger_perimetre(chemin):
    df = pd.read_csv(chemin, sep=None, engine="python", dtype=str)
    c_art = col(df, "Code article", "code_article", ou="scope")
    c_dec = col(df, "Code déclinaison", "Code declinaison", "code_declinaison",
                obligatoire=False)
    out = pd.DataFrame({"art": code(df[c_art])})
    out["dec"] = code(df[c_dec]) if c_dec else ""
    out["dec"] = out["dec"].fillna("")
    out = out.dropna(subset=["art"]).drop_duplicates()
    print("Scope            : %d items" % out["art"].nunique())
    return out


def charger_responsable(chemin, onglet="Base articles"):
    df = lire(chemin, onglet)
    ou = "price owner's workbook / " + onglet
    out = pd.DataFrame({
        "art": code(df[col(df, "Code article", ou=ou)]),
        "dec": code(df[col(df, "Code déclinaison", "Code declinaison", ou=ou)]).fillna(""),
        "frs": frs(df[col(df, "Fournisseur principal", "Fournisseur", ou=ou)]),
        "palier": palier(df[col(df, "Palier", ou=ou)]),
        "pa_resp": pd.to_numeric(df[col(df, "Nouveau PA", ou=ou)], errors="coerce"),
        "origine_resp": df[col(df, "Origine Nouveau PA", ou=ou)].astype(str),
        "fichier_resp": df[col(df, "Fichier source PA", obligatoire=False, ou=ou)]
        if col(df, "Fichier source PA", obligatoire=False) else None,
        "ref_frs_resp": df[col(df, "Réf. fournisseur", "Ref. fournisseur",
                               obligatoire=False, ou=ou)]
        if col(df, "Réf. fournisseur", "Ref. fournisseur", obligatoire=False) else None,
    })
    out = out.dropna(subset=["art"])
    print("Price owner      : %d rows, %d with a Nouveau PA"
          % (len(out), out["pa_resp"].notna().sum()))
    return out


def charger_notre(chemin, onglet=0):
    df = lire(chemin, onglet if onglet else 0)
    ou = "our proposals"
    c_pal = col(df, "Palier", "Qté palier", "Quantité palier", obligatoire=False)
    c_col = col(df, "colonne_prix", "Colonne prix", "Intitulé colonne prix",
                obligatoire=False)
    c_cle = col(df, "Clé de rapprochement", "Cle de rapprochement", obligatoire=False)
    c_sta = col(df, "Statut", "Motif", obligatoire=False)
    out = pd.DataFrame({
        "art": code(df[col(df, "Code article", ou=ou)]),
        "dec": code(df[col(df, "Code déclinaison", "Code declinaison",
                           obligatoire=False, ou=ou)]).fillna("")
        if col(df, "Code déclinaison", "Code declinaison", obligatoire=False) else "",
        "frs": frs(df[col(df, "Fournisseur principal", "Fournisseur", ou=ou)]),
        "palier": palier(df[c_pal]) if c_pal else np.nan,
        "pa_nous": pd.to_numeric(
            df[col(df, "Prix proposé", "Prix propose", "PA proposé", "Nouveau PA",
                   "Prix d'achat 2026", ou=ou)], errors="coerce"),
        "colonne_prix": df[c_col] if c_col else None,
        "cle_rappro": df[c_cle] if c_cle else None,
        "statut_nous": df[c_sta] if c_sta else None,
    })
    out = out.dropna(subset=["art"])
    print("Our proposals    : %d rows, %d with a price"
          % (len(out), out["pa_nous"].notna().sum()))
    return out


# --------------------------------------------------------------------------- reconciliation
def palier_unitaire_du_responsable(responsable):
    """Reduces the price owner's data to ONE row per item/variant/supplier.

    To be used only when our side carries no price break. Our proposals are
    UNIT prices, and a unit price can only be compared to a unit price. So we
    keep, in this order:

        1. the price-break-1 row, the per-unit price, the comparable one;
        2. failing that, the row WITHOUT a price break, which most often
           stands for the unit.

    Everything else is marked NON COMPARABLE. At Suppliers P, AP or CD, the
    price owner's smallest price break is 5, 6, sometimes 28: their price is
    volume-discounted there. Setting it against our unit price would produce a
    gap that says nothing about the price list, a FALSE GAP, in the opposite
    direction to the one on 13/08. Better a row classed "non comparable" than
    an invented divergence.
    """
    grp = ["art", "dec", "frs"]
    # Order of preference: break 1, then no break, then the rest by ascending
    # value. Only an explicit sort guarantees which one survives.
    rang = np.where(responsable["palier"] == 1, 0,
                    np.where(responsable["palier"].isna(), 1, 2))
    j = responsable.assign(_rang=rang).sort_values(["_rang", "palier"],
                                                   na_position="last")
    # How many price breaks does this item carry on the price owner's side?
    # Beyond one, the comparison only holds for the one we keep, and that has
    # to be said.
    j["Paliers chez le responsable"] = j.groupby(
        grp, dropna=False)["pa_resp"].transform("size")
    j = j.drop_duplicates(grp, keep="first")
    # A unit price either exists or it does not: it is this column that will
    # later decide whether we are entitled to speak of a gap at all.
    j["comparable_unitaire"] = j["_rang"] < 2
    return j.drop(columns="_rang").rename(
        columns={"palier": "Palier responsable retenu"})


def reconcilier(perim, responsable, notre):
    # Restriction to the frozen scope. We READ it, we do not recompute it.
    responsable = responsable.merge(perim, on=["art", "dec"], how="inner")
    notre = notre.merge(perim, on=["art", "dec"], how="inner")

    # The price owner's R4 key (item, variant, supplier, price break) assumes
    # that BOTH sides carry a price break. Our pa_etat_par_article.xlsx is a
    # PER-ITEM view: its price-break column is empty from end to end. Joining
    # on it only matched the 61 rows where the price owner was empty too, and
    # the script concluded "no divergence" after comparing almost nothing: an
    # unearned green light, more dangerous than a reported divergence.
    #
    # So we choose the key according to what the data actually carries, and we
    # say so on screen instead of assuming it.
    au_palier = "palier" in notre.columns and notre["palier"].notna().any()
    if au_palier:
        cle = ["art", "dec", "frs", "palier"]
        print("Comparison key: item + variant + supplier + price break")
    else:
        cle = ["art", "dec", "frs"]
        responsable = palier_unitaire_du_responsable(responsable)
        print("Comparison key: item + variant + supplier")
        print("  our prices are UNIT prices with no price break: we set them "
              "against the price owner's lowest break")
        multi = int((responsable["Paliers chez le responsable"] > 1).sum())
        if multi:
            print("  %d items carry several price breaks on the price owner's "
                  "side - the comparison only holds for the one kept, stated "
                  "in a column" % multi)

    m = responsable.merge(notre, on=cle, how="outer", indicator=True,
                          suffixes=("_j", "_n"))

    a_resp = m["pa_resp"].notna()
    a_nous = m["pa_nous"].notna()

    # In unit mode, the two prices can only be set against each other if the
    # price owner's really is a per-unit price. Otherwise the gap is not a
    # reading disagreement, it is volume discounting, and naming it otherwise
    # would send someone chasing a defect that does not exist.
    if "comparable_unitaire" in m.columns:
        # .astype(bool) is not cosmetic: the outer join introduces NaN, the
        # column falls back to dtype object, and "~True" is -2 there, a
        # non-zero integer, hence truthy. Without this cast every row becomes
        # "non comparable" without a single error being raised.
        opposable = m["comparable_unitaire"].fillna(True).astype(bool)
    else:
        opposable = pd.Series(True, index=m.index)

    m["Classe"] = np.select(
        [a_resp & a_nous & ~opposable,
         a_resp & a_nous & _egaux(m["pa_resp"], m["pa_nous"]),
         a_resp & a_nous,
         a_resp & ~a_nous,
         ~a_resp & a_nous],
        ["PALIER NON UNITAIRE", "ACCORD", "DIVERGENCE", "SEUL RESPONSABLE",
         "SEUL NOUS"],
        default="SANS PRIX")

    m["Rapport"] = (m["pa_nous"] / m["pa_resp"]).where(
        a_resp & a_nous & (m["pa_resp"] > 0))
    m["Écart %"] = 100 * (m["Rapport"] - 1)

    # Divergences split by the origin of the price owner's price. A divergence
    # against a supplier price list is a reading defect. A divergence against
    # the internal file is a commercial decision, and that is not the same job.
    org = m["origine_resp"].fillna("")
    m["Nature"] = np.where(
        m["Classe"] != "DIVERGENCE", "",
        np.where(org.str.startswith("Tarif frs") | (org == "Tarif fournisseur"),
                 "LECTURE (les deux lisent le tarif fournisseur)",
                 np.where(org.str.contains("INTERNE"),
                          "ARBITRAGE (le responsable retient le fichier interne)",
                          "ORIGINE INDÉTERMINÉE")))
    return m


def par_fournisseur(m):
    """Median ratio of ours to the price owner's per supplier, and per price
    column read.

    A median ratio that is stable and different from 1 over at least a few rows
    is not negotiated row by row: it is a column or a unit, and it is settled
    with a single decision for the whole supplier.
    """
    d = m[m["Classe"].isin(["ACCORD", "DIVERGENCE"])].copy()
    if d.empty:
        return pd.DataFrame()
    g = d.groupby("frs")
    out = pd.DataFrame({
        "Comparées": g.size(),
        "Accords": g.apply(lambda x: (x["Classe"] == "ACCORD").sum()),
        "Divergences": g.apply(lambda x: (x["Classe"] == "DIVERGENCE").sum()),
        "Rapport médian": g["Rapport"].median(),
        "Rapport min": g["Rapport"].min(),
        "Rapport max": g["Rapport"].max(),
    }).reset_index().rename(columns={"frs": "Fournisseur"})
    out["Dispersion"] = out["Rapport max"] - out["Rapport min"]

    def verdict(r):
        if r["Divergences"] == 0:
            return "OK"
        if r["Comparées"] < MIN_LIGNES_RATIO:
            return "TROP PEU DE LIGNES, vérifier à la main"
        if abs(r["Rapport médian"] - 1) <= TOLERANCE_RATIO:
            return "DÉSACCORDS PONCTUELS, voir ligne à ligne"
        if r["Dispersion"] <= 0.02:
            return ("SYSTÉMATIQUE x%.4f, une colonne ou une unité, un seul arbitrage"
                    % r["Rapport médian"])
        return "DISPERSÉ, lecture instable, reprendre le fournisseur"

    out["Verdict"] = out.apply(verdict, axis=1)
    return out.sort_values(["Divergences", "Comparées"], ascending=False)


def synthese(m, perim):
    n_art = perim["art"].nunique()
    c = m["Classe"].value_counts()
    couvert = m.loc[m["Classe"] != "SANS PRIX", "art"].nunique()
    lignes = [
        ("Articles du périmètre", n_art),
        ("Articles couverts par au moins un prix", couvert),
        ("Articles sans aucun prix", n_art - couvert),
        ("", ""),
        ("Lignes en accord", int(c.get("ACCORD", 0))),
        ("Lignes en divergence", int(c.get("DIVERGENCE", 0))),
        ("  dont défaut de lecture", int((m["Nature"].str.startswith("LECTURE")).sum())),
        ("  dont arbitrage interne", int((m["Nature"].str.startswith("ARBITRAGE")).sum())),
        ("Non comparable (le responsable n'a pas de prix unitaire)",
         int(c.get("PALIER NON UNITAIRE", 0))),
        ("Prix chez le responsable seulement", int(c.get("SEUL RESPONSABLE", 0))),
        ("Prix chez nous seulement", int(c.get("SEUL NOUS", 0))),
    ]
    return pd.DataFrame(lignes, columns=["Indicateur", "Valeur"])


# --------------------------------------------------------------------------- output
# "Palier responsable retenu" and "Paliers chez le responsable" only exist in
# the unit comparison mode. They say on WHICH of the price owner's rows the gap
# was computed: without them, a gap against an item with six price breaks would
# be unreadable.
COLS = ["art", "dec", "frs", "palier", "Palier responsable retenu",
        "Paliers chez le responsable",
        "pa_resp", "pa_nous", "Rapport", "Écart %",
        "origine_resp", "fichier_resp", "colonne_prix", "cle_rappro",
        "ref_frs_resp", "statut_nous", "Nature"]
NOMS = {"art": "Code article", "dec": "Code déclinaison", "frs": "Fournisseur",
        "palier": "Palier", "pa_resp": "PA responsable", "pa_nous": "PA proposé",
        "origine_resp": "Origine (responsable)",
        "fichier_resp": "Fichier source (responsable)",
        "colonne_prix": "Colonne prix lue (nous)", "cle_rappro": "Clé de rapprochement",
        "ref_frs_resp": "Réf. fournisseur", "statut_nous": "Statut (nous)"}


def _vue(m, classe):
    d = m[m["Classe"] == classe]
    d = d[[c for c in COLS if c in d.columns]].rename(columns=NOMS)
    return d.dropna(axis=1, how="all")


def ecrire(m, perim, chemin):
    os.makedirs(os.path.dirname(chemin) or ".", exist_ok=True)
    if os.path.exists(chemin):                      # the workbook may be open
        base, ext = os.path.splitext(chemin)
        try:
            os.remove(chemin)
        except PermissionError:
            chemin = "%s %s%s" % (base, pd.Timestamp.now().strftime("%Y-%m-%d %H%M"), ext)
    with pd.ExcelWriter(chemin, engine="openpyxl") as w:
        # sheet_name is a KEYWORD argument as of pandas 2: passing it
        # positionally raises a TypeError after the workbook has been opened.
        synthese(m, perim).to_excel(w, sheet_name="Synthese", index=False)
        par_fournisseur(m).to_excel(w, sheet_name="Par fournisseur", index=False)
        div = _vue(m, "DIVERGENCE")
        if "Écart %" in div.columns:
            div = div.sort_values("Écart %", key=lambda s: s.abs(),
                                  ascending=False)
        div.to_excel(w, sheet_name="Divergences", index=False)
        _vue(m, "SEUL RESPONSABLE").to_excel(w, sheet_name="Seul responsable",
                                             index=False)
        _vue(m, "SEUL NOUS").to_excel(w, sheet_name="Seul nous", index=False)
        # Neither an agreement nor a disagreement: the price owner only has
        # volume-discounted breaks on these items. To be picked up again when
        # our proposals carry a price break.
        _vue(m, "PALIER NON UNITAIRE").to_excel(
            w, sheet_name="Palier non unitaire", index=False)
        _vue(m, "SANS PRIX").to_excel(w, sheet_name="Sans prix", index=False)
    return chemin


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--perimetre", default=os.path.join("perimetre", "perimetre_v1_2.csv"))
    # The reference workbook is the price owner's BASE, the 30,197-row one
    # carrying the "Base articles" tab. Two traps avoided here:
    #   - the default folder did not exist on the workstation, and the old
    #     default could only fail;
    #   - "prix_valides.xlsx" has NO "Base articles" tab (its own is called
    #     Feuil1), so the default path and the default tab contradicted each
    #     other.
    # We aim at the SOURCE base, never at a "+ propositions" copy:
    # reconciling against a file into which our own prices have already been
    # poured would amount to comparing ourselves with ourselves and finding no
    # gap at all.
    p.add_argument("--responsable", default=os.path.join(
        ".", "data", "base_prix.xlsx"))
    p.add_argument("--onglet-responsable", default="Base articles")
    p.add_argument("--notre", default=os.path.join("sortie", "3-achats",
                                                   "pa_etat_par_article.xlsx"))
    p.add_argument("--sortie", default=os.path.join(
        "sortie", "3-achats", "reconciliation_responsable.xlsx"))
    a = p.parse_args()

    perim = charger_perimetre(a.perimetre)
    responsable = charger_responsable(a.responsable, a.onglet_responsable)
    notre = charger_notre(a.notre)

    m = reconcilier(perim, responsable, notre)
    print()
    print(synthese(m, perim).to_string(index=False))
    print()
    pf = par_fournisseur(m)
    if not pf.empty:
        suspects = pf[pf["Verdict"] != "OK"]
        if len(suspects):
            print("SUPPLIERS TO BE REVIEWED")
            print(suspects.to_string(index=False))
        else:
            print("No supplier in divergence.")
    chemin = ecrire(m, perim, a.sortie)
    print()
    print("Written: %s" % chemin)

    # A reading defect is not a nuance: as long as one remains, no further
    # proposal should go out to the price owner.
    lecture = int((m["Nature"].str.startswith("LECTURE")).sum())
    if lecture:
        print("!! %d READING divergence(s): to be settled before extending "
              "the workstream" % lecture)
        sys.exit(1)


if __name__ == "__main__":
    main()
