# -*- coding: utf-8 -*-
"""
PA workstream: propose a price through label matching.

    python pa_libelle.py --calibrer      measures the threshold on known cases
    python pa_libelle.py FOURNISSEUR_E   proposes, for one supplier
    python pa_libelle.py                 the whole unmatched scope

THE MOST DANGEROUS OF THE THREE WORKSTREAMS
    The price owner tested the EXACT label on his 514 unmatched articles:
    six recoveries.
    All of the gain is therefore in approximate matching, and approximate
    matching, on medical devices, gets the size, the side or the colour
    wrong without saying so.

WHAT IS REUSED
    `par_libelle.py` already does the hard part and paid dearly for it:
    normalisation, brand removal, stop words, and above all the
    DISCRIMINANTS. It is the module that learned that "MODELE EXEMPLE 1
    LIGHT" is not "EXTRA Light" and that twenty chairs of the GAMME
    EXEMPLE 1 range had inherited the code of a bariatric model.
    We keep its rule: a discriminant present on one side only REJECTS the
    candidate, it does not arbitrate between candidates.

WHAT THIS MODULE ADDS
    1. UNIQUENESS. The runner-up must be clearly behind: eight points of
       gap. Two candidates level with each other is not a coin toss, it is
       a non-match. `par_libelle` only purged ambiguity AFTERWARDS, when
       the same code landed on several articles; here we refuse to choose
       as soon as there is doubt.
    2. TOKEN SIMILARITY, on top of the share of our words found.
       `token_set_ratio` tolerates word order and labels of very different
       lengths, which is what catalogues do all the time.
    3. CORROBORATION BY THE PRICE. Never as a key: picking the line whose
       price most resembles the old PA would be circular, and would
       manufacture the FALSE GAP instead of detecting it. Only as a
       counter-check on a candidate already found through the label.
    4. CONFIDENCE LEVELS, and a review output.

WHAT NEVER LEAVES THIS MODULE
    A price. This module writes PROPOSALS into a review workbook, with the
    score, the price list label and the WMS label side by side. Nothing
    goes into pa_injectable.xlsx without human validation - the same rule
    as the near-miss references at the price owner's end: reported, never
    applied.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import par_libelle  # noqa: E402

try:
    from rapidfuzz import fuzz
except ImportError:  # pragma: no cover
    fuzz = None


# --- thresholds -------------------------------------------------------------
# The brief aimed at 92, "to be calibrated on known cases, not picked by
# feel". Calibration decided otherwise, and calibration is what counts.
#
# Measured on the articles already matched by exact reference, where the
# right answer is known (`--calibrer`):
#
#   threshold  FOURNISSEUR E            FOURNISSEUR A
#     84    18 right /  0 wrong     340 right / 14 wrong   precision 0.960
#     88    14 right /  0 wrong     335 right / 13 wrong   precision 0.963
#     92     0 right /  0 wrong     329 right / 13 wrong   precision 0.962
#
# At 92 NOTHING matches at Supplier E any more - precisely the supplier for
# which the label is the only route, since its price list carries no usable
# reference. 88 keeps its 14 matches without a single wrong one, and costs
# only five matches at Supplier A.
#
# Note above all that precision barely moves with the threshold (0.960 at
# 84, 0.962 at 96): the threshold is NOT the lever that removes errors. The
# remaining wrong matches score high - they are genuinely close products
# that no discriminant catches. Only review sees them.
SEUIL_SCORE = 88
ECART_UNICITE = 8

# Corroboration by the price: bounds taken from the brief.
CORROBORE = (0.80, 1.25)
PLAUSIBLE = (0.50, 2.00)


def modele(libelle, marque=None) -> str:
    """The label stripped of whatever tells variants apart.

    We match on the MODEL alone. Discriminants - size, side, colour,
    quantity - are not there to resemble, they are there to reject:
    "MODELE EXEMPLE 2 T3 DROIT" and "MODELE EXEMPLE 2 T4 DROIT" are 95%
    identical and
    are not the same product. Mixing them into the score would amount to
    rewarding resemblance where identity must be required.
    """
    normalise = par_libelle._normaliser(libelle)
    prefixe = par_libelle._normaliser(marque) if marque else ""
    if prefixe and normalise.startswith(prefixe):
        normalise = normalise[len(prefixe):]
    garde = [mot for mot in normalise.split()
             if mot not in par_libelle.DISCRIMINANTS
             and mot not in par_libelle.MOTS_VIDES
             and not mot.isdigit()]
    # ATTEMPT REJECTED - do not try again without measuring. Truncating each
    # word to five characters (par_libelle.RACINE) to reconcile "CEINT" and
    # "CEINTURE" looks obvious: the two catalogues abbreviate the same
    # product differently. Measured on the articles already matched by
    # reference, it is in fact WORSE - precision falls from 1.000 to 0.805
    # at Supplier E and from 0.960 to 0.955 at Supplier A. Truncation wins
    # a few matches and invents more. We keep whole words.
    return " ".join(garde)


def score_modele(a: str, b: str) -> float:
    """Token similarity between two models, from 0 to 100."""
    if not a or not b:
        return 0.0
    if fuzz is None:
        # Fallback without rapidfuzz: share of common tokens. Cruder, but
        # the module has to stay usable on a machine without the library.
        ja, jb = set(a.split()), set(b.split())
        return 100.0 * len(ja & jb) / max(len(ja | jb), 1)
    return float(fuzz.token_set_ratio(a, b))


def discriminants_compatibles(libelle_wms, libelle_tarif, marque=None) -> bool:
    """Discriminants must be IDENTICAL, not similar.

    Present on one side and absent on the other: the candidate is rejected,
    not arbitrated. This is the rule that costs the most recall and avoids
    the most errors.
    """
    d_wms = par_libelle.discriminants(libelle_wms, marque)
    d_tarif = par_libelle.discriminants(libelle_tarif)
    if d_wms ^ d_tarif:
        return False
    n_wms = par_libelle.nombres(libelle_wms)
    n_tarif = par_libelle.nombres(libelle_tarif)
    if n_wms and n_tarif and not (n_wms & n_tarif):
        return False
    q_wms = par_libelle.quantite(libelle_wms)
    q_tarif = par_libelle.quantite(libelle_tarif)
    if q_wms and q_tarif and not (q_wms & q_tarif):
        return False
    return True


def meilleur_candidat(libelle_wms, lignes: list[dict], marque=None,
                      seuil=SEUIL_SCORE, ecart=ECART_UNICITE) -> dict | None:
    """The single acceptable candidate, or nothing, with the reason why.

    Three cumulative conditions: score above the threshold, uniqueness, and
    identical discriminants.
    """
    mon_modele = modele(libelle_wms, marque)
    if len(mon_modele.split()) < 2:
        return {"retenu": None, "motif": "libellé du WMS trop pauvre pour "
                                         "rapprocher (moins de deux mots)"}

    notes = []
    for ligne in lignes:
        if not discriminants_compatibles(libelle_wms, ligne["libelle"], marque):
            continue
        note = score_modele(mon_modele, ligne.get("modele")
                            or modele(ligne["libelle"]))
        notes.append((note, ligne))
    if not notes:
        return {"retenu": None,
                "motif": "aucun candidat aux discriminants identiques"}

    notes.sort(key=lambda x: -x[0])
    premier, ligne = notes[0]
    second = notes[1][0] if len(notes) > 1 else 0.0

    if premier < seuil:
        return {"retenu": None,
                "motif": f"meilleur score {premier:.0f} < seuil {seuil}"}
    if (premier - second) < ecart:
        # Lesson of the FALSE GAP fixed on 13/08: at a tie, we do not draw
        # lots.
        return {"retenu": None,
                "motif": f"candidats trop proches ({premier:.0f} contre "
                         f"{second:.0f}) — ambiguïté, pas de rapprochement"}
    return {"retenu": ligne, "score": round(premier, 1),
            "second": round(second, 1), "motif": ""}


def corroborer_par_prix(prix_candidat, dpa) -> tuple[str, str]:
    """Does the candidate found by the label survive its ratio to the DPA?

    The price never CHOOSES the candidate - that would be circular, and the
    gap would be zero by construction. It only confirms or contradicts a
    candidate already found some other way.
    """
    if prix_candidat is None or pd.isna(prix_candidat):
        return "", "pas de prix au candidat"
    if dpa is None or pd.isna(dpa) or dpa <= 0:
        # The guard rail only works when an old price exists. On an article
        # that was never bought, no check is possible, and the candidate
        # must not come out with the same confidence as a corroborated one.
        return "PROPOSÉ", "aucun DPA : corroboration impossible"
    rapport = float(prix_candidat) / float(dpa)
    if CORROBORE[0] <= rapport <= CORROBORE[1]:
        return "À RELIRE", f"corroboré par le prix (×{rapport:.2f})"
    if PLAUSIBLE[0] <= rapport <= PLAUSIBLE[1]:
        return "À RELIRE", (f"plausible (×{rapport:.2f}) — relecture "
                            f"obligatoire")
    # Beyond that, a pack divisor may still explain it.
    try:
        import pa_conditionnement
        facteur = pa_conditionnement.entier_rond_le_plus_proche(rapport)
    except Exception:
        facteur = None
    if facteur:
        return "À RELIRE", (f"rapport ×{rapport:.2f} expliqué par un "
                            f"conditionnement de {facteur} — à vérifier")
    return "", f"rejeté : rapport ×{rapport:.2f} hors de toute explication"


# --------------------------------------------------------------------------
# Calibration


def calibrer(motif: str, seuils=(84, 88, 90, 92, 94, 96)) -> pd.DataFrame:
    """Measures the threshold on cases whose answer is already known.

    Articles matched by EXACT reference provide a free test set: we know
    which price list line is the right one. We hide the reference from
    them, leave them only the label, and see whether matching finds the
    same line again. It is the only way to pick a threshold other than by
    feel.
    """
    import pa_completer as P

    articles = P.articles_du_perimetre(motif)
    if articles.empty:
        print("no article for this supplier")
        return pd.DataFrame()
    noms = tuple(sorted(articles["Nom fournisseur"].dropna().unique()))
    tarif = P.tarif_du_fournisseur(motif, noms)
    if tarif.empty or "designation" not in tarif.columns:
        print("no usable price list")
        return pd.DataFrame()

    from extracteurs.base import cle_ref_stricte
    tarif = tarif.dropna(subset=["designation"]).copy()
    tarif["cle"] = tarif["ref_fournisseur"].map(cle_ref_stricte)
    lignes = [{"libelle": r["designation"], "cle": r["cle"],
               "modele": modele(r["designation"])}
              for _, r in tarif.iterrows()]

    articles = articles.copy()
    articles["cle"] = articles["ref_fournisseur_wms"].map(cle_ref_stricte)
    cles_tarif = set(tarif["cle"].dropna())
    connus = articles[articles["cle"].isin(cles_tarif)]
    print(f"test set: {len(connus)} articles matched by reference")
    if connus.empty:
        return pd.DataFrame()

    resultats = []
    for seuil in seuils:
        justes = faux = muets = 0
        for _, art in connus.iterrows():
            res = meilleur_candidat(art["Libellé déclinaison ^(1)"], lignes,
                                    art["Nom fournisseur"], seuil=seuil)
            if not res or not res.get("retenu"):
                muets += 1
                continue
            if res["retenu"]["cle"] == art["cle"]:
                justes += 1
            else:
                faux += 1
        total = justes + faux
        resultats.append({
            "Seuil": seuil,
            "Retrouvés": justes,
            "Faux": faux,
            "Sans réponse": muets,
            "Précision": round(justes / total, 3) if total else None,
            "Rappel": round(justes / len(connus), 3),
        })
    return pd.DataFrame(resultats)


def proposer(motif: str) -> pd.DataFrame:
    """One supplier's candidates, for review - no price applied.

    We only work on the articles that none of the five keys matched and
    whose supplier does have a price list: anywhere else, the label has
    nothing to add that the keys have not already given, more reliably.
    """
    import pa_completer as P
    from extracteurs.base import cle_ref_stricte

    articles = P.articles_du_perimetre(motif)
    if articles.empty:
        return pd.DataFrame()
    noms = tuple(sorted(articles["Nom fournisseur"].dropna().unique()))
    tarif = P.tarif_du_fournisseur(motif, noms)
    if tarif.empty or "designation" not in tarif.columns:
        return pd.DataFrame()

    tarif = tarif.dropna(subset=["designation"]).copy()
    tarif["cle"] = tarif["ref_fournisseur"].map(cle_ref_stricte)
    lignes = []
    for _, r in tarif.iterrows():
        lignes.append({
            "libelle": r["designation"],
            "modele": modele(r["designation"]),
            "ref": r.get("ref_fournisseur"),
            "prix": r.get("prix_achat_unitaire_ht"),
            "fichier": r.get("fichier_tarif"),
        })

    # Whatever no key matched
    articles = articles.copy()
    articles["cle"] = articles["ref_fournisseur_wms"].map(cle_ref_stricte)
    articles["cle_fab"] = articles["ref_fabricant"].map(cle_ref_stricte)
    cles = set(tarif["cle"].dropna())
    reste = articles[~articles["cle"].isin(cles) & ~articles["cle_fab"].isin(cles)]
    if reste.empty:
        return pd.DataFrame()

    pa = P.pa_actuels().set_index("code_article")["pa_wms"]
    sorties = []
    for _, art in reste.iterrows():
        res = meilleur_candidat(art["Libellé déclinaison ^(1)"], lignes,
                                art["Nom fournisseur"])
        if not res:
            continue
        code = str(art["Code article"]).strip()
        dpa = pa.get(code)
        if not res.get("retenu"):
            sorties.append({
                "Code article": code,
                "Libellé WMS": art["Libellé déclinaison ^(1)"],
                "Fournisseur": art["Nom fournisseur"],
                "Niveau": "", "Motif": res["motif"], "Ancien PA": dpa,
            })
            continue
        ligne = res["retenu"]
        niveau, verdict = corroborer_par_prix(ligne["prix"], dpa)
        sorties.append({
            "Code article": code,
            "Libellé WMS": art["Libellé déclinaison ^(1)"],
            "Libellé tarif": ligne["libelle"],
            "Fournisseur": art["Nom fournisseur"],
            "Réf. tarif": ligne["ref"],
            "Prix proposé": ligne["prix"],
            "Ancien PA": dpa,
            "Score": res["score"],
            "2e candidat": res["second"],
            "Niveau": niveau,
            "Motif": verdict,
            "Fichier source": ligne["fichier"],
            "VENTE_12M": art.get("VENTE_12M"),
        })
    return pd.DataFrame(sorties)


def ecrire_relecture(tables: list[pd.DataFrame]) -> Path | None:
    """The review workbook, sorted so that human time pays off."""
    import mise_en_forme
    from main import SORTIE_ACHATS

    utiles = [t for t in tables if not t.empty]
    if not utiles:
        print("no candidate")
        return None
    tout = pd.concat(utiles, ignore_index=True)
    retenus = tout[tout["Niveau"].isin(("À RELIRE", "PROPOSÉ"))].copy()
    rejets = tout[~tout["Niveau"].isin(("À RELIRE", "PROPOSÉ"))]

    if not retenus.empty:
        # One price list label proposed for SEVERAL of our articles is a
        # generic label, not a match. "Canne anglaise" gets assigned to four
        # crutches of the same GAMME EXEMPLE 2 range in different colours:
        # at most one is
        # right, and nothing says which. This is the lesson of
        # par_libelle.purger_doublons, applied here upstream - we do not
        # purge silently, we say it in a column so that review sees it.
        partages = retenus["Libellé tarif"].value_counts()
        retenus["Candidat partagé"] = retenus["Libellé tarif"].map(partages)
        trop = retenus["Candidat partagé"] > 1
        retenus.loc[trop, "Motif"] = (
            retenus.loc[trop, "Motif"] + " — ATTENTION : ce libellé de tarif est "
            "proposé à " + retenus.loc[trop, "Candidat partagé"].astype(str)
            + " articles, il est trop générique pour les distinguer")
        # A shared candidate falls back to the lowest level, whatever its
        # corroboration by the price.
        retenus.loc[trop, "Niveau"] = "PROPOSÉ"

        # Review time goes first where it pays: what sells, then what costs
        # a lot. Reliable candidates first.
        retenus["_sur"] = retenus["Niveau"].eq("À RELIRE")
        retenus["_vend"] = retenus["VENTE_12M"].fillna(False).astype(bool)
        retenus = retenus.sort_values(
            ["_sur", "_vend", "Prix proposé"],
            ascending=[False, False, False]).drop(columns=["_sur", "_vend"])

    chemin = mise_en_forme.chemin_ecriture(SORTIE_ACHATS / "pa_libelle_relecture.xlsx")
    with pd.ExcelWriter(chemin, engine="openpyxl") as w:
        retenus.to_excel(w, sheet_name="À relire", index=False, startrow=3)
        rejets.to_excel(w, sheet_name="Sans candidat", index=False, startrow=3)
    mise_en_forme.formater(chemin, {
        "À relire": {
            "ligne_entete": 4, "figer_colonne": 3,
            "titre": "Candidats trouvés par le LIBELLÉ — à relire un par un",
            "sous_titre": "Aucun de ces prix n'est injectable en l'état. "
                          "Mesuré : 96 % de justesse chez le Fournisseur A, donc "
                          "environ un sur vingt-cinq est faux."},
        "Sans candidat": {
            "ligne_entete": 4, "figer_colonne": 3,
            "titre": "Aucun candidat retenu",
            "sous_titre": "Le motif dit pourquoi : score trop bas, ambiguïté, "
                          "ou discriminants différents"},
    })
    print(f"  {len(retenus)} to review, {len(rejets)} without a candidate")
    print(f"  -> {chemin}")
    return chemin


def main() -> None:
    args = [a for a in sys.argv[1:]]
    if fuzz is None:
        print("! rapidfuzz missing: falling back to a common-token score")
    if "--calibrer" in args:
        args.remove("--calibrer")
        motif = args[0] if args else "FOURNISSEUR_E"
        print(f"Threshold calibration on {motif}\n")
        table = calibrer(motif)
        if not table.empty:
            pd.set_option("display.width", 200)
            print(table.to_string(index=False))
            print()
            print("How to read this: \"Faux\" is the only figure that costs")
            print("money. A wrong, silent price is worse than no price.")
        return
    cibles = args or ["FOURNISSEUR_E", "FOURNISSEUR_A", "FOURNISSEUR_S",
                      "FOURNISSEUR_Q", "FOURNISSEUR_AQ"]
    print(f"Label matching - threshold {SEUIL_SCORE}, "
          f"uniqueness gap {ECART_UNICITE}\n")
    tables = []
    for motif in cibles:
        print(f"--- {motif}")
        try:
            tables.append(proposer(motif))
        except Exception as err:
            print(f"    ! {type(err).__name__} : {err}")
    ecrire_relecture(tables)
    print()
    print("Nothing is injected. These candidates await human review.")


if __name__ == "__main__":
    main()
