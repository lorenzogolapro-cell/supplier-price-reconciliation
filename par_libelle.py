# -*- coding: utf-8 -*-
"""
Finds a product by its label when its reference is not enough.

The supplier reference is authoritative, but it is not always right in the
WMS: at Supplier S, the chairs of the GAMME EXEMPLE 1 range carry a
reference that does not follow the manufacturer's, while their label is
reliable. So the label serves two purposes:

  - RECOVER   when no reference matches, look the product up by its name,
              within the same supplier;
  - CHALLENGE when a reference does match but the two labels have nothing
              in common, treat it with suspicion.

The method is the one proven on VIDAL: we measure the share of OUR words
found at the supplier's end, not mutual resemblance, because catalogues
describe things at greater length than our labels do, and we require the
numbers to agree. That last point is the essential guard rail: without it,
a size L would be given the code of a size S.

A code obtained this way is marked "libellé": it is a solid lead, not a
certainty, and it must not go into the WMS without review.
"""

from __future__ import annotations

import re
import unicodedata

import pandas as pd

# Vocabulary too common to tell two products apart
MOTS_VIDES = {
    "DE", "DU", "LA", "LE", "LES", "ET", "AVEC", "SANS", "POUR", "PAR",
    "EN", "UN", "UNE", "DES", "SUR", "AU", "AUX",
    "CM", "MM", "ML", "KG", "CL", "GR",
    "BTE", "BOITE", "SACHET", "CARTON", "PCS", "PIECE", "PIECES", "UNITE",
    "LOT", "PAIRE", "TAILLE", "POINTURE", "COLORIS", "REF", "MODELE",
}

# Stem length: absorbs plurals and our own truncated labels
RACINE = 5

# Share of our words that must be found at the supplier's end
SEUIL = 0.6

# Every number, even glued to letters: "T39", "CH12", "0,8cm"
NOMBRE = re.compile(r"\d+(?:[.,]\d+)?")

# Words that tell apart two products of the same range. Present on one side
# only, they disqualify the match - which is exactly what was missing when
# "MODELE EXEMPLE 1 LIGHT" was given the code of "MODELE EXEMPLE 1 EXTRA
# Light", and when twenty chairs of the GAMME EXEMPLE 1 range inherited the
# code of a bariatric model.
DISCRIMINANTS = {
    # intensity / range
    "EXTRA", "SUPER", "ULTRA", "MAXI", "MINI", "PLUS", "PREMIUM", "LIGHT",
    "NORMAL", "STANDARD", "CONFORT", "BASIC", "PRO", "EVO", "BARIATRIQUE",
    # morphology
    "XXL", "XXS", "JUNIOR", "ENFANT", "ADULTE", "PEDIATRIQUE", "BEBE",
    "COURT", "LONG", "LARGE", "ETROIT", "HAUT", "BAS",
    # one- or two-letter sizes: too short to be kept as words, yet they are
    # what separates the S from the L
    "S", "M", "L", "XL", "XS", "TS", "TM", "TL", "TXL", "TU",
    # side and shape
    "GAUCHE", "DROITE", "DROIT", "PLIANT", "FIXE", "ELECTRIQUE", "MANUEL",
    # colours: two shades are two articles
    "NOIR", "NOIRE", "BLANC", "BLANCHE", "BLEU", "BLEUE", "ROUGE", "VERT",
    "VERTE", "GRIS", "GRISE", "BEIGE", "TAUPE", "MARRON", "ROSE", "JAUNE",
    "ORANGE", "VIOLET", "ANTHRACITE", "CHOCO", "GREGE", "CHINE", "SAPHIR",
}

# Quantity per pack: "SACHET DE 24", "BTE/30", "LOT DE 5", "x 12". Two
# different packs are two articles - as in the pack of 24 that was given
# the code of the pack of 14.
QUANTITE = re.compile(
    r"(?:SACHET|BOITE|BTE|LOT|PAQUET|PACK|CARTON|BLISTER|SET|POCHE)"
    r"[\s/]*(?:DE\s*)?(\d{1,4})\b|"
    r"\b(?:X|PAR)\s*(\d{1,4})\b|"
    r"\b(\d{1,4})\s*(?:PCS|PIECES?|UNITES?|U\.)\b",
    re.IGNORECASE,
)


def _normaliser(texte) -> str:
    if texte is None or (isinstance(texte, float) and texte != texte):
        return ""
    sans_accent = unicodedata.normalize("NFKD", str(texte).upper())
    sans_accent = "".join(c for c in sans_accent
                          if not unicodedata.combining(c))
    # The apostrophe of a contraction does not separate two words: "MODEL'S"
    # counts as ONE word, not "MODEL" + "S". Removing it BEFORE replacing
    # the rest with spaces avoids manufacturing a spurious "S", which also
    # happens to be the abbreviation for size Small in DISCRIMINANTS.
    # Without this removal, two "MODEL'S ..." products ALWAYS SHARE that
    # "S", and the discriminant check (colour, side) no longer sees their
    # real disagreements: "MODEL'S GO NOIR" was passing as compatible with
    # "Model's Go Beige", found on 17/09 in the Supplier AN batch.
    sans_apostrophe = re.sub(r"['’]", "", sans_accent)
    return re.sub(r"[^A-Z0-9]+", " ", sans_apostrophe)


def mots(texte, marque: str | None = None) -> set:
    """Meaningful words of a label, brand and common vocabulary removed.

    Our labels start with the brand ("FOURNISSEUR S GAMME EXEMPLE 1...")
    which the manufacturer's catalogue does not repeat: counting it would
    skew the score in both directions.
    """
    normalise = _normaliser(texte)
    prefixe = _normaliser(marque)
    if prefixe and normalise.startswith(prefixe):
        normalise = normalise[len(prefixe):]

    retenus = set()
    for mot in normalise.split():
        if len(mot) < 3 or mot in MOTS_VIDES or mot.isdigit():
            continue
        retenus.add(mot[:RACINE])
    return retenus


def nombres(texte) -> set:
    """Numbers in the label: sizes, calibres, dimensions.

    Two products whose numbers differ are not the same one, however much
    their words resemble each other.
    """
    valeurs = set()
    for brut in NOMBRE.findall(str(texte or "")):
        normalise = brut.replace(",", ".").rstrip("0").rstrip(".")
        valeurs.add(normalise.lstrip("0") or "0")
    return valeurs


def discriminants(texte, marque: str | None = None) -> set:
    """Words that tell apart two products of the same range."""
    normalise = _normaliser(texte)
    prefixe = _normaliser(marque)
    if prefixe and normalise.startswith(prefixe):
        normalise = normalise[len(prefixe):]
    return {mot for mot in normalise.split() if mot in DISCRIMINANTS}


def quantite(texte) -> set:
    """Pack quantities quoted in the label."""
    valeurs = set()
    for groupes in QUANTITE.findall(str(texte or "")):
        for valeur in groupes:
            if valeur:
                valeurs.add(str(int(valeur)))
    return valeurs


def indexer(catalogue: pd.DataFrame, colonne_fournisseur: str = "fournisseur",
            colonne_libelle: str = "designation") -> dict:
    """Prepares the catalogue for lookup: one index per supplier."""
    index: dict[str, list] = {}
    for ligne in catalogue.itertuples():
        fournisseur = getattr(ligne, colonne_fournisseur, None)
        libelle = getattr(ligne, colonne_libelle, None)
        ean = getattr(ligne, "ean", None)
        if not fournisseur or not libelle or not ean:
            continue
        index.setdefault(str(fournisseur).upper(), []).append({
            "libelle": libelle,
            "ean": ean,
            "reference": getattr(ligne, "ref_fournisseur", None),
            "mots": mots(libelle),
            "nombres": nombres(libelle),
            "discriminants": discriminants(libelle),
            "quantite": quantite(libelle),
        })
    return index


def chercher(libelle: str, candidats: list[dict], marque: str | None = None,
             seuil: float = SEUIL) -> dict | None:
    """Best catalogue product for one of our labels."""
    nos_mots = mots(libelle, marque)
    nos_nombres = nombres(libelle)
    nos_discriminants = discriminants(libelle, marque)
    notre_quantite = quantite(libelle)
    if len(nos_mots) < 2:
        return None

    meilleur, score_max = None, 0.0
    for candidat in candidats:
        if not candidat["mots"]:
            continue
        # Numbers come first: they tell the variants apart
        if nos_nombres and candidat["nombres"] and not (
            nos_nombres & candidat["nombres"]
        ):
            continue

        # A discriminant word present on one side only separates two
        # products of the same range: "Extra Light" is not "Light", and a
        # bariatric model is not the standard one.
        if nos_discriminants ^ candidat["discriminants"]:
            continue

        # Two different packs are two articles: the pack of 24 does not
        # carry the code of the pack of 14.
        if notre_quantite and candidat["quantite"] and not (
            notre_quantite & candidat["quantite"]
        ):
            continue

        communs = nos_mots & candidat["mots"]
        if not communs:
            continue
        score = len(communs) / len(nos_mots)
        if score > score_max or (
            score == score_max and meilleur is not None
            and len(candidat["libelle"]) < len(meilleur["libelle"])
        ):
            meilleur, score_max = candidat, score

    if meilleur is None or score_max < seuil:
        return None
    return {**meilleur, "score": round(score_max, 2)}


def enrichir(df: pd.DataFrame, catalogue: pd.DataFrame,
             colonne_fournisseur: str = "nom_fournisseur_wms",
             colonne_libelle: str = "designation_wms",
             correspondance=None) -> pd.DataFrame:
    """Fills in missing EANs by label matching.

    `correspondance` maps the supplier name as the WMS carries it to the
    name the catalogue publishes its lines under (the WMS says
    "EXEMPLE' S.A.", the price list says "EXEMPLE SA").
    """
    df = df.copy()
    if "source_ean" not in df.columns:
        df["source_ean"] = None

    index = indexer(catalogue)
    a_chercher = df["ean"].isna()
    if not a_chercher.any():
        return df

    trouves = 0
    for position in df.index[a_chercher]:
        fournisseur = df.at[position, colonne_fournisseur]
        if not fournisseur or pd.isna(fournisseur):
            continue

        cible = correspondance(fournisseur) if correspondance else fournisseur
        candidats = index.get(str(cible).upper())
        if not candidats:
            continue

        resultat = chercher(df.at[position, colonne_libelle], candidats,
                            marque=fournisseur)
        if not resultat:
            continue
        df.at[position, "ean"] = resultat["ean"]
        df.at[position, "source_ean"] = "libellé"
        df.at[position, "libelle_fournisseur"] = resultat["libelle"]
        # The score says how much of our label was found at the supplier's
        # end: above 0.8 the product is almost always the right one, below
        # that it needs a look.
        df.at[position, "concordance_libelle"] = resultat["score"]
        trouves += 1

    print(f"  EANs recovered from the label: {trouves} (to be checked)")
    return purger_doublons(df)


def purger_doublons(df: pd.DataFrame, sources: tuple = ("libellé",)
                    ) -> pd.DataFrame:
    """Removes inferred codes assigned to SEVERAL of our articles.

    An EAN code designates one product and one only. When a label match
    gives the same code to twenty chairs that differ only in colour and
    seat width, nineteen are wrong - and nothing says which one is right.
    We remove them all: a gap can be filled later, whereas a wrong code
    makes people scan one product for another.

    Codes coming from a price list or a reference are not concerned: a
    supplier may legitimately publish the same code on two lines of its
    catalogue.
    """
    deduits = df["source_ean"].isin(sources) & df["ean"].notna()
    if not deduits.any():
        return df

    comptes = df.loc[deduits, "ean"].value_counts()
    partages = set(comptes[comptes > 1].index)
    if not partages:
        return df

    a_purger = deduits & df["ean"].isin(partages)
    retires = int(a_purger.sum())
    df.loc[a_purger, ["ean", "source_ean", "libelle_fournisseur",
                      "concordance_libelle"]] = None
    print(f"  {retires} removed: {len(partages)} codes assigned to "
          f"several articles (match too loose)")
    return df
