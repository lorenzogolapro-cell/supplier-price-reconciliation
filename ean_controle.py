# -*- coding: utf-8 -*-
"""
Quatre controles sur chaque code EAN retenu, du moins cher au plus cher.

    python ean_controle.py            tout le classeur
    python ean_controle.py --perimetre  le perimetre fige seulement

Sortie : sortie/2-chantier-ean/ean_controle.xlsx

Ce script NE SUPPRIME RIEN et n'ecrit pas dans « EAN distributeur - Vn.xlsx ».
Il pose un verdict en colonne. Le tri se fait a la sortie, dans ce qu'on
lit et dans ce qu'on emporte — jamais dans le journal de la chaine.

Les quatre controles
--------------------
0. STRUCTURE   la cle de controle, la longueur, et les prefixes qui ne
               designent pas un produit
1. UNICITE     un code designe UN article. Quand plusieurs le portent,
               leurs libelles disent s'il s'agit de vrais doublons du
               WMS ou de declinaisons distinctes
2. LOT         le prefixe GS1 doit etre celui du lot (fournisseur x
               source). C'est le test qui tient a 100 % sur les tarifs
3. CROISEMENT  deux sources qui documentent le meme article doivent dire
               la meme chose

Le niveau 4 — la confrontation au libelle chez EUDAMED — vit dans
`ean_verifier_libelle.py` : il coute une minute par code et ne peut pas
tourner ici. Le niveau 5, la douchette, ne s'automatise pas.

Ce que les mesures du 14/09 ont montre, et qui dicte ces controles
-------------------------------------------------------------------
UNICITE. « MOTORISATION FOURNISSEUR GB MODELE-C » porte onze coloris sous
une seule reference fournisseur, donc un seul code. La chaine les
etiquette « articles WMS en doublon » et les garde. Ce ne sont pas des
doublons : ce sont onze articles, dont dix portent un code qui ne les
designe pas. 515 lignes portent ce motif. Dix-neuf autres portent le bon
diagnostic, « GTIN de gamme », pour exactement le meme phenomene.

Le code lui-meme est souvent JUSTE au niveau du modele — son prefixe GS1
est allemand, et ce fabricant l'est aussi. C'est son affectation a
l'article qui est fausse. D'ou un motif distinct : on ne jette pas un
code de gamme, on dit qu'il est pose trop bas.

LOT. Le prefixe GS1 des codes lus dans un tarif concorde a 100 % avec
celui du fournisseur ; ceux venus d'EUDAMED, a 27,8 %. Le prefixe est
donc un test, et le lot (fournisseur x source) sa bonne maille : une
contradiction dans un lot condamne le lot, pas seulement la ligne.

Le biais a connaitre, et qu'on ne doit PAS compter comme une erreur :
EUDAMED indexe le FABRICANT. Un fournisseur qui distribue sans fabriquer
portera legitimement un prefixe etranger — on achete du fournisseur B
chez un distributeur. Le controle le dit, il ne condamne pas : un lot entier
qui decroche est probablement un distributeur, une ligne isolee dans un
lot sain est probablement une erreur.

CROISEMENT. Sur les 39 articles documentes a la fois par EUDAMED et par
une autre source, EUDAMED en contredit 31 — dont 5 contre un tarif. Deux
methodes sans rapport, le prefixe et le croisement, donnent le meme
verdict : c'est ce qui rend la conclusion solide.
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

# Les sources dont le prefixe fait reference. Mesure du 14/09 :
# tarif 100 %, fournisseur 98 %, entrepot lu sur le produit.
SOURCES_DE_REFERENCE = {"tarif", "fournisseur", "entrepôt"}

# Longueur du prefixe entreprise retenu pour le test de lot.
#
# Six chiffres depuis le 16/09, et non sept. Sept coupait le fournisseur
# U en 7654320 / 7654321 / 7654322 — trois sous-plages de la MEME
# entreprise, comptees comme etrangeres les unes aux autres : 33 + 25 + 9
# codes qui s'ignoraient, la ou six les reunit en 67.
#
# Mesure sur V16 avant de trancher, les deux effets et pas seulement
# celui qui arrange :
#   concordants   5 548 -> 5 610   (84,4 % -> 85,3 % des 6 575 testables)
#   lots rompus      31 -> 29      sur 81 lots testables
#   prefixes partages par plusieurs fournisseurs — le risque de
#   confusion que raccourcir fait courir :  231 -> 237
#
# Un gain modeste pour six confusions de plus. On prend, mais il ne faut
# pas descendre plus bas : a cinq chiffres le prefixe ne designe plus une
# entreprise, et le test cesserait de vouloir dire quelque chose.
PREFIXE = 6

# Un lot en dessous de ce taux est repute rompu. On ne descend pas a
# 100 % : un fournisseur legitime peut porter deux prefixes (rachat,
# filiale etrangere, sous-traitance).
TAUX_LOT = 0.90

# En deca, un lot est trop petit pour qu'un taux veuille dire quelque
# chose.
LOT_MINIMUM = 5

# Prefixes qui ne designent pas un produit du commerce.
#   02x et 2xx : usage interne a une entreprise (balances, codes maison)
#   977 / 978 / 979 : presse et livres
#   980 a 99x : recus de detaxe et bons de reduction
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
        raise SystemExit("Aucun fichier « EAN distributeur - Vn.xlsx ».")
    return fichiers[-1]


def _normaliser(serie: pd.Series) -> pd.Series:
    """Le piege pandas : un code article devient « 32818.0 » des qu'une
    valeur manque dans la colonne, et la jointure tombe a zero sans lever
    la moindre erreur."""
    return (serie.astype(str).str.strip()
            .str.replace(r"\.0$", "", regex=True)
            .replace({"nan": pd.NA, "": pd.NA}))


# ---------------------------------------------------------------- niveau 0

def niveau0_structure(df: pd.DataFrame) -> pd.Series:
    """La forme du code, sans rien savoir du produit."""
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
    """Ces libelles designent-ils des articles DIFFERENTS ?

    On ne compare pas les chaines — « BLEU CLAIR » et « BLEU FONCE » se
    ressemblent beaucoup. On compare ce qui distingue deux produits d'une
    meme gamme : les mots discriminants (couleur, cote, intensite) et les
    nombres (taille, pointure, calibre).
    """
    signatures = {
        (frozenset(par_libelle.discriminants(libelle)),
         frozenset(par_libelle.nombres(libelle)))
        for libelle in libelles
    }
    return len(signatures) > 1


def niveau1_unicite(df: pd.DataFrame) -> pd.Series:
    """Un code, un article — sauf vrai doublon du WMS."""
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
            motifs.loc[lot.index] = ""  # vrais doublons du WMS : acceptable
    return motifs


# ---------------------------------------------------------------- niveau 2

def niveau2_lot(df: pd.DataFrame) -> tuple[pd.Series, pd.DataFrame]:
    """Le prefixe GS1 doit etre celui du lot (fournisseur x source)."""
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
    """Un fichier de source, quel que soit l'etage de son en-tete."""
    if not chemin.exists():
        print(f"    {nom} : fichier absent, croisement impossible")
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
            print(f"    {nom} : {len(sortie)} codes")
            return sortie.drop_duplicates("k")
    print(f"    {nom} : colonnes introuvables, ignoré")
    return pd.DataFrame(columns=["k", "ean", "source"])


def niveau3_croisement(df: pd.DataFrame) -> pd.Series:
    """Deux sources qui documentent le meme article doivent concorder."""
    print("  sources croisées :")
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
    print(f"{len(df)} codes exploitables ({fichier.name})")

    if perimetre_seul:
        references = perimetre_liste.references()
        codes = perimetre_liste.codes_article()
        dedans = (df.get("Référence", pd.Series("", index=df.index))
                  .fillna("").str.strip().isin(references)
                  | df["Code article"].fillna("").str.strip().isin(codes))
        df = df[dedans].copy()
        print(f"  {perimetre_liste.entete()}")
        print(f"  {len(df)} dans le périmètre figé")

    df = df.reset_index(drop=True)

    # Le niveau 2 rend aussi sa table de lots : il est appele a part.
    motifs = {}
    for nom, controle in NIVEAUX[:2]:
        motifs[nom] = controle(df)
        print(f"  contrôle {nom} : {int(motifs[nom].ne('').sum())} en défaut")
    motifs["2 — lot"], lots = niveau2_lot(df)
    print(f"  contrôle 2 — lot : {int(motifs['2 — lot'].ne('').sum())} "
          f"en défaut, sur {len(lots)} lots")
    motifs["3 — croisement"] = niveau3_croisement(df)
    print(f"  contrôle 3 — croisement : "
          f"{int(motifs['3 — croisement'].ne('').sum())} en défaut")

    # Le premier controle en defaut donne le verdict : les niveaux sont
    # ordonnes du plus structurel au plus circonstanciel, et un code dont
    # la cle est fausse n'a pas besoin qu'on lui cherche un prefixe.
    ordre = ["0 — structure", "1 — unicité", "2 — lot", "3 — croisement"]
    df["Contrôle en défaut"] = ""
    df["Détail"] = ""
    for niveau in reversed(ordre):
        touche = motifs[niveau].ne("")
        df.loc[touche, "Contrôle en défaut"] = niveau
        df.loc[touche, "Détail"] = motifs[niveau][touche]
    df["Verdict"] = df["Contrôle en défaut"].map(
        lambda x: "à vérifier" if x else "passe les 4 contrôles")

    # Combien de controles chaque code echoue : un code qui en echoue
    # deux n'est pas dans la meme situation qu'un code qui en echoue un.
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
    print(f"\n{en_defaut} codes en défaut sur {len(df)}")
    print(f"→ {chemin.name}")


if __name__ == "__main__":
    main()
