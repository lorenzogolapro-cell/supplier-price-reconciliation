# -*- coding: utf-8 -*-
"""
Four checks on every EAN code retained, from the cheapest to the dearest.

    python ean_controle.py            the whole workbook
    python ean_controle.py --perimetre  the frozen scope only

Output: sortie/2-chantier-ean/ean_controle.xlsx

This script DELETES NOTHING and writes nothing into
"EAN distributeur - Vn.xlsx". It records a verdict in a column. The
sorting happens on the way out, in what gets read and what gets taken to
the warehouse, never in the pipeline's own log.

The four checks
---------------
0. STRUCTURE   the check digit, the length, and the prefixes that do not
               designate a product
1. UNICITE     one code designates ONE article. When several of them
               carry it, their labels say whether these are genuine WMS
               duplicates or distinct variants
2. LOT         the GS1 prefix must be the batch's (supplier x source).
               This is the test that holds at 100 % on price lists
3. CROISEMENT  two sources documenting the same article must say the
               same thing

Level 4, the comparison against the label in EUDAMED, lives in
`ean_verifier_libelle.py`: it costs a minute per code and cannot run
here. Level 5, the hand scanner, cannot be automated.

What the 14/09 measurements showed, and what dictates these checks
------------------------------------------------------------------
UNICITE. "MOTORISATION FOURNISSEUR GB MODELE-C" carries eleven colours
under a single supplier reference, hence a single code. The pipeline
labels them "articles WMS en doublon" and keeps them. They are not
duplicates: they are eleven articles, ten of which carry a code that
does not designate them. 515 rows carry that reason. Nineteen others
carry the right diagnosis, "GTIN de gamme", for exactly the same
phenomenon.

The code itself is often RIGHT at model level: its GS1 prefix is German,
and so is the manufacturer. What is wrong is its assignment to the
article. Hence a distinct reason: we do not throw away a range code, we
say it has been placed one level too low.

LOT. The GS1 prefix of the codes read in a price list agrees 100 % with
the supplier's own; the ones coming from EUDAMED, 27.8 %. The prefix is
therefore a test, and the batch (supplier x source) its right
granularity: one contradiction inside a batch condemns the batch, not
only the row.

The bias to know about, and which must NOT be counted as an error:
EUDAMED indexes the MANUFACTURER. A supplier that distributes without
manufacturing will legitimately carry a foreign prefix: we buy supplier
B's goods from a distributor. The check states it, it does not condemn:
a whole batch breaking away is probably a distributor, an isolated row
inside a healthy batch is probably an error.

CROISEMENT. Of the 39 articles documented both by EUDAMED and by another
source, EUDAMED contradicts 31, including 5 against a price list. Two
unrelated methods, the prefix and the cross-check, return the same
verdict: that is what makes the conclusion solid.
"""

from __future__ import annotations

import sys
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import pandas as pd  # noqa: E402

import mise_en_forme  # noqa: E402
import par_libelle  # noqa: E402
import perimetre_liste  # noqa: E402
from extracteurs.base import chemin_lisible, ean_est_valide  # noqa: E402

SORTIE_CHANTIER = RACINE / "sortie" / "2-chantier-ean"
SORTIE = SORTIE_CHANTIER / "ean_controle.xlsx"

# The sources whose prefix is authoritative. Measured on 14/09:
# price list 100 %, supplier 98 %, warehouse read off the product itself.
SOURCES_DE_REFERENCE = {"tarif", "fournisseur", "entrepôt"}

# Length of the company prefix used for the batch test.
#
# Six digits since 16/09, and not seven. Seven split supplier U into
# 7654320 / 7654321 / 7654322, three sub-ranges of the SAME company,
# counted as foreign to one another: 33 + 25 + 9 codes ignoring each
# other, where six digits reunite them into 67.
#
# Measured on V16 before deciding, both effects and not just the
# convenient one:
#   agreeing       5 548 -> 5 610  (84.4 % -> 85.3 % of the 6 575 testable)
#   broken batches    31 -> 29     out of 81 testable batches
#   prefixes shared by several suppliers, the risk of confusion that
#   shortening creates:  231 -> 237
#
# A modest gain for six more confusions. We take it, but we must not go
# any lower: at five digits the prefix no longer designates a company,
# and the test would stop meaning anything.
PREFIXE = 6

# A batch below this rate is deemed broken. We do not go up to 100 %: a
# legitimate supplier can carry two prefixes (acquisition, foreign
# subsidiary, subcontracting).
TAUX_LOT = 0.90

# Below this, a batch is too small for a rate to mean anything.
LOT_MINIMUM = 5

# Prefixes that do not designate a retail product.
#   02x and 2xx: internal company use (scales, in-house codes)
#   977 / 978 / 979: press and books
#   980 to 99x: tax refund receipts and coupons
def _prefixe_interdit(code: str) -> str:
    tete = code[:3]
    if tete.startswith("02") or tete.startswith("2"):
        return "préfixe d'usage interne (02x/2xx) : jamais un code produit"
    if tete in {"977", "978", "979"}:
        return "préfixe presse ou livre (977/978/979)"
    if tete.startswith("98") or tete.startswith("99"):
        return "préfixe coupon ou détaxe (98x/99x)"
    return ""


COLONNES = ["Code article", "Référence", "Désignation", "Fournisseur",
            "Réf. fournisseur", "Code EAN", "Source", "Fiabilité",
            "Motif du partage", "Verdict", "Contrôle en défaut", "Détail"]


def _derniere_version() -> Path:
    import re
    fichiers = sorted(
        SORTIE_CHANTIER.glob("EAN distributeur - V*.xlsx"),
        key=lambda p: int(re.search(r"- V(\d+)", p.name).group(1)),
    )
    if not fichiers:
        raise SystemExit("No \"EAN distributeur - Vn.xlsx\" file found.")
    return fichiers[-1]


def _normaliser(serie: pd.Series) -> pd.Series:
    """The pandas trap: an article code becomes "32818.0" as soon as one
    value is missing from the column, and the join drops to zero without
    raising the slightest error."""
    return (serie.astype(str).str.strip()
            .str.replace(r"\.0$", "", regex=True)
            .replace({"nan": pd.NA, "": pd.NA}))


# ---------------------------------------------------------------- niveau 0

def niveau0_structure(df: pd.DataFrame) -> pd.Series:
    """The shape of the code, knowing nothing about the product."""
    motifs = []
    for code in df["Code EAN"].fillna(""):
        code = str(code).strip()
        if not code:
            motifs.append("aucun code")
        elif not code.isdigit():
            motifs.append("caractères non numériques")
        elif len(code) == 8:
            motifs.append("EAN-8 : les 58 déjà retirés venaient tous "
                          "d'une référence prise pour un code")
        elif len(code) != 13:
            motifs.append(f"longueur {len(code)}, attendu 13")
        elif not ean_est_valide(code):
            motifs.append("clé de contrôle fausse")
        else:
            motifs.append(_prefixe_interdit(code))
    return pd.Series(motifs, index=df.index)


# ---------------------------------------------------------------- niveau 1

def _declinaisons_distinctes(libelles: list[str]) -> bool:
    """Do these labels designate DIFFERENT articles?

    We do not compare the strings: "BLEU CLAIR" and "BLEU FONCE" look
    very much alike. We compare what tells two products of the same range
    apart: the discriminating words (colour, side, strength) and the
    numbers (size, shoe size, gauge).
    """
    signatures = {
        (frozenset(par_libelle.discriminants(libelle)),
         frozenset(par_libelle.nombres(libelle)))
        for libelle in libelles
    }
    return len(signatures) > 1


def niveau1_unicite(df: pd.DataFrame) -> pd.Series:
    """One code, one article, unless it is a genuine WMS duplicate."""
    motifs = pd.Series("", index=df.index)
    for code, lot in df.groupby(df["Code EAN"].fillna("")):
        if not code or len(lot) < 2:
            continue
        libelles = [str(x) for x in lot["Désignation"].fillna("")]
        if _declinaisons_distinctes(libelles):
            motifs.loc[lot.index] = (
                f"{len(lot)} articles distincts sous un seul code : "
                f"code de gamme posé au niveau de l'article")
        else:
            motifs.loc[lot.index] = ""  # genuine WMS duplicates: acceptable
    return motifs


# ---------------------------------------------------------------- niveau 2

def niveau2_lot(df: pd.DataFrame) -> tuple[pd.Series, pd.DataFrame]:
    """The GS1 prefix must be the batch's (supplier x source)."""
    travail = df.copy()
    travail["prefixe"] = travail["Code EAN"].fillna("").str[:PREFIXE]
    travail["src"] = travail["Source"].fillna("").str.strip()

    reference = travail[travail["src"].isin(SOURCES_DE_REFERENCE)]
    connus = reference.groupby("Fournisseur")["prefixe"].apply(set).to_dict()

    testable = travail["Fournisseur"].isin(connus) & travail["prefixe"].ne("")
    concordant = travail.apply(
        lambda r: r["prefixe"] in connus.get(r["Fournisseur"], set()), axis=1)

    lots = (travail[testable].assign(concordant=concordant[testable])
            .groupby(["Fournisseur", "src"])["concordant"]
            .agg(codes="size", concordants="sum").reset_index())
    lots["taux"] = (lots["concordants"] / lots["codes"]).round(3)
    lots["verdict du lot"] = lots.apply(
        lambda r: "lot sain" if r["taux"] >= TAUX_LOT
        else ("lot rompu" if r["codes"] >= LOT_MINIMUM
              else "lot trop petit pour conclure"), axis=1)

    rompus = {(r["Fournisseur"], r["src"]) for _, r in lots.iterrows()
              if r["verdict du lot"] == "lot rompu"}

    motifs = pd.Series("", index=df.index)
    for position in df.index[testable & ~concordant]:
        fournisseur = travail.at[position, "Fournisseur"]
        source = travail.at[position, "src"]
        if (fournisseur, source) in rompus:
            ligne = lots[(lots["Fournisseur"] == fournisseur)
                         & (lots["src"] == source)].iloc[0]
            motifs.at[position] = (
                f"préfixe étranger au lot {fournisseur} × {source}, "
                f"rompu ({ligne['concordants']}/{ligne['codes']} concordants)")
        else:
            motifs.at[position] = (
                "préfixe étranger au fournisseur, dans un lot par ailleurs "
                "sain — erreur isolée ou produit d'un autre fabricant")
    return motifs, lots.sort_values(["taux", "codes"])


# ---------------------------------------------------------------- niveau 3

def _lire_source(chemin: Path, feuille, col_article, col_ean,
                 nom: str) -> pd.DataFrame:
    """One source file, whatever row its header sits on."""
    if not chemin.exists():
        print(f"    {nom}: file missing, cross-check impossible")
        return pd.DataFrame(columns=["k", "ean", "source"])
    for entete in (3, 0):
        try:
            df = pd.read_excel(chemin_lisible(chemin), sheet_name=feuille,
                               header=entete, dtype=str)
        except Exception:
            continue
        if col_article in df.columns and col_ean in df.columns:
            sortie = pd.DataFrame({
                "k": _normaliser(df[col_article]),
                "ean": df[col_ean].astype(str).str.strip(),
                "source": nom,
            }).dropna(subset=["k"])
            sortie = sortie[sortie["ean"].str.fullmatch(r"\d{8,14}")]
            print(f"    {nom}: {len(sortie)} codes")
            return sortie.drop_duplicates("k")
    print(f"    {nom}: columns not found, skipped")
    return pd.DataFrame(columns=["k", "ean", "source"])


def niveau3_croisement(df: pd.DataFrame) -> pd.Series:
    """Two sources documenting the same article must agree."""
    print("  cross-checked sources:")
    sources = pd.concat([
        _lire_source(SORTIE_CHANTIER / "ean_eudamed_reference.xlsx",
                     "Codes EAN", "Code article", "Code EAN", "eudamed"),
        _lire_source(SORTIE_CHANTIER / "ean_fournisseurs.xlsx",
                     "Codes EAN", "Code article", "Code EAN", "fournisseur"),
        _lire_source(SORTIE_CHANTIER / "ean_vidal.xlsx",
                     "Rapprochements", "code_article", "ean", "vidal"),
        _lire_source(SORTIE_CHANTIER / "catalogue_wms_complete.xlsx",
                     "Articles", "code_article", "ean", "tarif"),
    ], ignore_index=True)

    par_article: dict[str, list[tuple[str, str]]] = {}
    for ligne in sources.itertuples():
        par_article.setdefault(ligne.k, []).append((ligne.source, ligne.ean))

    cles = _normaliser(df["Code article"])
    motifs = pd.Series("", index=df.index)
    for position, cle in cles.items():
        retenu = str(df.at[position, "Code EAN"] or "").strip()
        if pd.isna(cle) or not retenu:
            continue
        source_retenue = str(df.at[position, "Source"] or "").strip()
        contredisent = [
            source for source, code in par_article.get(cle, [])
            if source != source_retenue and code and code != retenu
        ]
        if contredisent:
            motifs.at[position] = (
                "contredit par " + ", ".join(sorted(set(contredisent))))
    return motifs


# ----------------------------------------------------------------- sortie

NIVEAUX = [
    ("0 — structure", niveau0_structure),
    ("1 — unicité", niveau1_unicite),
    ("3 — croisement", niveau3_croisement),
]


def main() -> None:
    perimetre_seul = "--perimetre" in sys.argv
    fichier = _derniere_version()
    df = pd.read_excel(chemin_lisible(fichier), sheet_name="Codes EAN",
                       skiprows=3, dtype=str)
    df = df[df["Exploitable"].fillna("").str.strip() == "oui"].copy()
    print(f"{len(df)} usable codes ({fichier.name})")

    if perimetre_seul:
        references = perimetre_liste.references()
        codes = perimetre_liste.codes_article()
        dedans = (df.get("Référence", pd.Series("", index=df.index))
                  .fillna("").str.strip().isin(references)
                  | df["Code article"].fillna("").str.strip().isin(codes))
        df = df[dedans].copy()
        print(f"  {perimetre_liste.entete()}")
        print(f"  {len(df)} within the frozen scope")

    df = df.reset_index(drop=True)

    # Level 2 also returns its batch table: it is called separately.
    motifs = {}
    for nom, controle in NIVEAUX[:2]:
        motifs[nom] = controle(df)
        print(f"  check {nom}: {int(motifs[nom].ne('').sum())} failing")
    motifs["2 — lot"], lots = niveau2_lot(df)
    print(f"  check 2 (lot): {int(motifs['2 — lot'].ne('').sum())} "
          f"failing, over {len(lots)} batches")
    motifs["3 — croisement"] = niveau3_croisement(df)
    print(f"  check 3 (croisement): "
          f"{int(motifs['3 — croisement'].ne('').sum())} failing")

    # The first failing check gives the verdict: the levels are ordered
    # from the most structural to the most circumstantial, and a code
    # whose check digit is wrong needs no prefix hunt.
    ordre = ["0 — structure", "1 — unicité", "2 — lot", "3 — croisement"]
    df["Contrôle en défaut"] = ""
    df["Détail"] = ""
    for niveau in reversed(ordre):
        touche = motifs[niveau].ne("")
        df.loc[touche, "Contrôle en défaut"] = niveau
        df.loc[touche, "Détail"] = motifs[niveau][touche]
    df["Verdict"] = df["Contrôle en défaut"].map(
        lambda x: "à vérifier" if x else "passe les 4 contrôles")

    # How many checks each code fails: a code failing two is not in the
    # same situation as a code failing one.
    df["Contrôles échoués"] = sum(motifs[n].ne("").astype(int)
                                  for n in ordre)

    synthese = (df.groupby(["Source", "Contrôle en défaut"]).size()
                .rename("Codes").reset_index()
                .sort_values(["Contrôle en défaut", "Codes"],
                             ascending=[True, False]))

    sortie = df[[c for c in COLONNES if c in df.columns]
                + ["Contrôles échoués"]]
    sortie = sortie.sort_values(["Contrôle en défaut", "Source",
                                 "Fournisseur"], ascending=[False, True, True])

    SORTIE_CHANTIER.mkdir(parents=True, exist_ok=True)
    chemin = mise_en_forme.chemin_ecriture(SORTIE)
    with pd.ExcelWriter(chemin, engine="openpyxl") as writer:
        sortie.to_excel(writer, sheet_name="Contrôles", index=False,
                        startrow=3)
        lots.to_excel(writer, sheet_name="Lots", index=False, startrow=3)
        synthese.to_excel(writer, sheet_name="Synthèse", index=False,
                          startrow=3)

    en_defaut = int(df["Contrôle en défaut"].ne("").sum())
    mise_en_forme.formater(chemin, options={
        "Contrôles": {
            "ligne_entete": 4,
            "titre": "Contrôle des codes EAN retenus",
            "sous_titre": (
                f"{len(df)} codes | {en_defaut} en défaut sur au moins un "
                f"contrôle | rien n'est supprimé, rien n'est écrit dans "
                f"« {fichier.name} »"),
        },
        "Lots": {
            "ligne_entete": 4,
            "titre": "Concordance du préfixe GS1 par lot (fournisseur × source)",
            "sous_titre": (
                f"un lot sous {TAUX_LOT:.0%} est rompu | un lot entier qui "
                f"décroche est souvent un distributeur, pas une erreur"),
        },
        "Synthèse": {
            "ligne_entete": 4,
            "titre": "Quel contrôle échoue, et pour quelle source",
            "sous_titre": "un code peut échouer à plusieurs contrôles",
        },
    })
    print(f"\n{en_defaut} codes failing out of {len(df)}")
    print(f"→ {chemin.name}")


if __name__ == "__main__":
    main()
