# -*- coding: utf-8 -*-
"""
Chantier PA — proposer un prix par rapprochement de libellé.

    python pa_libelle.py --calibrer      mesure le seuil sur des cas connus
    python pa_libelle.py FOURNISSEUR_E   propose, pour un fournisseur
    python pa_libelle.py                 tout le périmètre non rapproché

LE PLUS DANGEREUX DES TROIS CHANTIERS
    Le responsable des prix a testé le libellé EXACT sur ses 514 non
    rapprochés : six récupérations.
    Tout le gain est donc dans l'approximatif — et l'approximatif, sur du
    matériel médical, se trompe de taille, de côté ou de coloris sans le dire.

CE QUI EST RÉUTILISÉ
    `par_libelle.py` fait déjà le plus dur et l'a payé cher : normalisation,
    retrait de la marque, mots vides, et surtout les DISCRIMINANTS — c'est lui
    qui a appris que « MODELE EXEMPLE 1 LIGHT » n'est pas « EXTRA Light » et
    que vingt fauteuils de la GAMME EXEMPLE 1 avaient hérité du code d'un
    modèle bariatrique.
    On garde sa règle : un discriminant présent d'un seul côté REJETTE le
    candidat, il ne l'arbitre pas.

CE QUE CE MODULE AJOUTE
    1. UNICITÉ. Le deuxième meilleur candidat doit être nettement derrière —
       huit points d'écart. Deux candidats à égalité, ce n'est pas un tirage au
       sort, c'est un non-rapprochement. `par_libelle` ne purgeait l'ambiguïté
       qu'APRÈS coup, quand un même code atterrissait sur plusieurs articles ;
       ici on refuse de choisir dès qu'il y a doute.
    2. SIMILARITÉ PAR JETONS, en plus de la part de nos mots retrouvée.
       `token_set_ratio` supporte l'ordre des mots et les libellés de longueurs
       très différentes, ce que les catalogues font tout le temps.
    3. CORROBORATION PAR LE PRIX. Jamais comme clé — choisir la ligne dont le
       prix ressemble le plus à l'ancien PA serait circulaire, et fabriquerait
       le FAUX ÉCART au lieu de le détecter. Uniquement comme contre-épreuve
       d'un candidat déjà trouvé par le libellé.
    4. NIVEAUX DE CONFIANCE, et une sortie de relecture.

CE QUI NE SORT JAMAIS D'ICI
    Un prix. Ce module écrit des PROPOSITIONS dans un classeur de relecture,
    avec le score, le libellé du tarif et le libellé du WMS en regard. Rien ne
    part dans pa_injectable.xlsx sans validation humaine — même règle que les
    références proches chez le responsable des prix : signalées, jamais
    appliquées.
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


# --- seuils -----------------------------------------------------------------
# Le brief visait 92 « à calibrer sur des cas connus, pas à choisir au jugé ».
# La calibration a tranché autrement, et c'est elle qui fait foi.
#
# Mesuré sur les articles déjà rapprochés par référence exacte, où la bonne
# réponse est connue (`--calibrer`) :
#
#   seuil      FOURNISSEUR E            FOURNISSEUR A
#     84    18 justes /  0 faux     340 justes / 14 faux   précision 0,960
#     88    14 justes /  0 faux     335 justes / 13 faux   précision 0,963
#     92     0 justes /  0 faux     329 justes / 13 faux   précision 0,962
#
# 92 ne rapproche plus RIEN chez le Fournisseur E — précisément le fournisseur
# pour lequel le libellé est la seule voie, son tarif ne portant pas de
# référence exploitable. 88 garde ses 14 appariements sans un seul faux, et ne
# coûte que cinq appariements chez le Fournisseur A.
#
# Noter surtout que la précision ne bouge quasiment pas avec le seuil (0,960 à
# 84, 0,962 à 96) : le seuil n'est PAS le levier qui élimine les erreurs. Les
# faux restants ont un score élevé — ce sont des produits réellement proches
# dont aucun discriminant ne capte la différence. Seule la relecture les voit.
SEUIL_SCORE = 88
ECART_UNICITE = 8

# Corroboration par le prix : bornes du brief.
CORROBORE = (0.80, 1.25)
PLAUSIBLE = (0.50, 2.00)


def modele(libelle, marque=None) -> str:
    """Le libellé débarrassé de ce qui distingue les déclinaisons.

    On rapproche sur le MODÈLE seul. Les discriminants — taille, côté,
    coloris, quantité — ne servent pas à ressembler, ils servent à rejeter :
    « MODELE EXEMPLE 2 T3 DROIT » et « MODELE EXEMPLE 2 T4 DROIT » sont
    identiques à 95 % et
    ne sont pas le même produit. Les mêler au score reviendrait à récompenser
    la ressemblance là où il faut exiger l'identité.
    """
    normalise = par_libelle._normaliser(libelle)
    prefixe = par_libelle._normaliser(marque) if marque else ""
    if prefixe and normalise.startswith(prefixe):
        normalise = normalise[len(prefixe):]
    garde = [mot for mot in normalise.split()
             if mot not in par_libelle.DISCRIMINANTS
             and mot not in par_libelle.MOTS_VIDES
             and not mot.isdigit()]
    # ESSAI REJETÉ — ne pas refaire sans mesurer. Tronquer chaque mot à cinq
    # caractères (par_libelle.RACINE) pour réconcilier « CEINT » et
    # « CEINTURE » paraît évident : les deux catalogues abrègent différemment
    # le même produit. Mesuré sur les articles déjà rapprochés par référence,
    # c'est pourtant PIRE — la précision tombe de 1,000 à 0,805 chez le
    # Fournisseur E et de 0,960 à 0,955 chez le Fournisseur A. La troncature gagne
    # quelques appariements et en invente davantage. On garde les mots entiers.
    return " ".join(garde)


def score_modele(a: str, b: str) -> float:
    """Similarité par jetons entre deux modèles, de 0 à 100."""
    if not a or not b:
        return 0.0
    if fuzz is None:
        # Repli sans rapidfuzz : part des jetons communs. Moins fin, mais le
        # module doit rester utilisable sur un poste sans la bibliothèque.
        ja, jb = set(a.split()), set(b.split())
        return 100.0 * len(ja & jb) / max(len(ja | jb), 1)
    return float(fuzz.token_set_ratio(a, b))


def discriminants_compatibles(libelle_wms, libelle_tarif, marque=None) -> bool:
    """Les discriminants doivent être IDENTIQUES, pas ressemblants.

    Présent d'un côté et absent de l'autre : le candidat est rejeté, pas
    arbitré. C'est la règle qui coûte le plus de rappel et qui évite le plus
    d'erreurs.
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
    """Le seul candidat acceptable, ou rien — avec le motif du refus.

    Trois conditions cumulatives : score au-delà du seuil, unicité, et
    discriminants identiques.
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
        # Leçon du FAUX ÉCART corrigé le 13/08 : à égalité, on ne tire pas au sort.
        return {"retenu": None,
                "motif": f"candidats trop proches ({premier:.0f} contre "
                         f"{second:.0f}) — ambiguïté, pas de rapprochement"}
    return {"retenu": ligne, "score": round(premier, 1),
            "second": round(second, 1), "motif": ""}


def corroborer_par_prix(prix_candidat, dpa) -> tuple[str, str]:
    """Le candidat trouvé par le libellé résiste-t-il au rapport avec le DPA ?

    Le prix ne CHOISIT jamais le candidat — ce serait circulaire, et l'écart
    serait nul par construction. Il ne fait que confirmer ou infirmer un
    candidat déjà trouvé autrement.
    """
    if prix_candidat is None or pd.isna(prix_candidat):
        return "", "pas de prix au candidat"
    if dpa is None or pd.isna(dpa) or dpa <= 0:
        # Le garde-fou ne fonctionne que s'il existe un ancien prix. Sur un
        # article jamais acheté, aucun contrôle n'est possible, et le candidat
        # ne doit pas ressortir avec la même confiance qu'un candidat corroboré.
        return "PROPOSÉ", "aucun DPA : corroboration impossible"
    rapport = float(prix_candidat) / float(dpa)
    if CORROBORE[0] <= rapport <= CORROBORE[1]:
        return "À RELIRE", f"corroboré par le prix (×{rapport:.2f})"
    if PLAUSIBLE[0] <= rapport <= PLAUSIBLE[1]:
        return "À RELIRE", (f"plausible (×{rapport:.2f}) — relecture "
                            f"obligatoire")
    # Au-delà, un diviseur de conditionnement peut encore l'expliquer.
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
    """Mesure le seuil sur des cas dont on connaît déjà la réponse.

    Les articles rapprochés par RÉFÉRENCE exacte donnent un jeu d'épreuve
    gratuit : on sait quelle ligne du tarif est la bonne. On leur cache la
    référence, on ne leur laisse que le libellé, et on regarde si le
    rapprochement retrouve la même ligne. C'est la seule façon de choisir un
    seuil autrement qu'au jugé.
    """
    import pa_completer as P

    articles = P.articles_du_perimetre(motif)
    if articles.empty:
        print("aucun article pour ce fournisseur")
        return pd.DataFrame()
    noms = tuple(sorted(articles["Nom fournisseur"].dropna().unique()))
    tarif = P.tarif_du_fournisseur(motif, noms)
    if tarif.empty or "designation" not in tarif.columns:
        print("aucun tarif exploitable")
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
    print(f"jeu d'épreuve : {len(connus)} articles rapprochés par référence")
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
    """Les candidats d'un fournisseur, pour relecture — aucun prix appliqué.

    On ne travaille que sur les articles qu'aucune des cinq clés n'a rapprochés
    et dont le fournisseur a bien un tarif : ailleurs, le libellé n'a rien à
    apporter que les clés n'aient déjà donné, plus sûrement.
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

    # Ce qu'aucune clé n'a rapproché
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
    """Le classeur de relecture, trié pour que le temps humain paie."""
    import mise_en_forme
    from main import SORTIE_ACHATS

    utiles = [t for t in tables if not t.empty]
    if not utiles:
        print("aucun candidat")
        return None
    tout = pd.concat(utiles, ignore_index=True)
    retenus = tout[tout["Niveau"].isin(("À RELIRE", "PROPOSÉ"))].copy()
    rejets = tout[~tout["Niveau"].isin(("À RELIRE", "PROPOSÉ"))]

    if not retenus.empty:
        # Un même libellé de tarif proposé à PLUSIEURS de nos articles est un
        # libellé générique, pas une correspondance. « Canne anglaise » se voit
        # attribuer quatre cannes d'une même GAMME EXEMPLE 2 de coloris
        # différents : au plus une est
        # la bonne, et rien ne dit laquelle. C'est la leçon de
        # par_libelle.purger_doublons, appliquée ici en amont — on ne purge pas
        # en silence, on le dit en colonne pour que la relecture le voie.
        partages = retenus["Libellé tarif"].value_counts()
        retenus["Candidat partagé"] = retenus["Libellé tarif"].map(partages)
        trop = retenus["Candidat partagé"] > 1
        retenus.loc[trop, "Motif"] = (
            retenus.loc[trop, "Motif"] + " — ATTENTION : ce libellé de tarif est "
            "proposé à " + retenus.loc[trop, "Candidat partagé"].astype(str)
            + " articles, il est trop générique pour les distinguer")
        # Un candidat partagé retombe au niveau le plus bas, quelle que soit sa
        # corroboration par le prix.
        retenus.loc[trop, "Niveau"] = "PROPOSÉ"

        # Le temps de relecture va d'abord là où il rapporte : ce qui se vend,
        # puis ce qui coûte cher. Les candidats fiables d'abord.
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
    print(f"  {len(retenus)} à relire, {len(rejets)} sans candidat")
    print(f"  -> {chemin}")
    return chemin


def main() -> None:
    args = [a for a in sys.argv[1:]]
    if fuzz is None:
        print("! rapidfuzz absent : score de repli par jetons communs")
    if "--calibrer" in args:
        args.remove("--calibrer")
        motif = args[0] if args else "FOURNISSEUR_E"
        print(f"Calibration du seuil sur {motif}\n")
        table = calibrer(motif)
        if not table.empty:
            pd.set_option("display.width", 200)
            print(table.to_string(index=False))
            print()
            print("Lire : « Faux » est le seul chiffre qui coûte cher. Un prix")
            print("faux et silencieux est pire qu'un article sans prix.")
        return
    cibles = args or ["FOURNISSEUR_E", "FOURNISSEUR_A", "FOURNISSEUR_S",
                      "FOURNISSEUR_Q", "FOURNISSEUR_AQ"]
    print(f"Rapprochement par libellé — seuil {SEUIL_SCORE}, "
          f"écart d'unicité {ECART_UNICITE}\n")
    tables = []
    for motif in cibles:
        print(f"--- {motif}")
        try:
            tables.append(proposer(motif))
        except Exception as err:
            print(f"    ! {type(err).__name__} : {err}")
    ecrire_relecture(tables)
    print()
    print("Rien n'est injecté. Ces candidats attendent une relecture humaine.")


if __name__ == "__main__":
    main()
