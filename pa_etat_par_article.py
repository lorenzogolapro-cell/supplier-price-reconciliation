# -*- coding: utf-8 -*-
"""
Etat du prix d'achat 2026, article par article, sur tout le perimetre.

    python pa_etat_par_article.py

Les autres classeurs du chantier montrent chacun un morceau : un
fournisseur, une sous-population, les propositions. Aucun ne repond a la
question qu'on pose vraiment devant un article : celui-la, ou en est-il,
et si rien n'a ete fait, pourquoi ?

Une ligne par article du perimetre fige, les 4 731, y compris ceux dont
il n'y a rien a dire — c'est justement leur nombre qui compte. Aucun
filtre n'est applique : tout ce qui pourrait en etre un devient une
colonne, pour qu'on puisse trier sans avoir a relancer un traitement.

Sortie : sortie/3-achats/pa_etat_par_article.xlsx
"""

from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import mise_en_forme  # noqa: E402
import perimetre_liste  # noqa: E402
from extracteurs.base import chemin_lisible  # noqa: E402
from main import SORTIE_ACHATS  # noqa: E402

DONNEES = Path("./data")
import wms_extract  # noqa: E402
EXTRACT = wms_extract.chemin()

REPRISES = ["Ancien PA", "Nouveau PA", "Écart %", "Origine Nouveau PA",
            # La reference telle qu'elle figure DANS LE TARIF : c'est
            # celle avec laquelle on commande, et elle ne coincide pas
            # toujours avec celle de la fiche article.
            "Réf. fournisseur",
            "Fichier source PA", "Alerte", "Clé de rapprochement"]


def code_article(colonne: pd.Series) -> pd.Series:
    """Le code article en texte, sans le « .0 » que pandas y colle.

    Une colonne d'entiers qui contient un vide devient flottante, et
    « 32818 » se lit alors « 32818.0 ». La jointure tombe a zero sans
    lever la moindre erreur — c'est le genre de panne qui se voit
    seulement au resultat final, quand il est trop tard.
    """
    texte = colonne.astype(str).str.strip()
    return texte.str.replace(r"\.0$", "", regex=True)


def classeur_fusionne() -> Path:
    """Le dernier classeur prix validés + propositions produit."""
    trouves = sorted(DONNEES.glob("prix_valides + propositions *.xlsx"),
                     key=lambda f: f.stat().st_mtime)
    if not trouves:
        raise SystemExit("Lancer d'abord pa_fusionner.py")
    return trouves[-1]


def prix_connus() -> pd.DataFrame:
    """Ce que le classeur consolide sait, par code article.

    On lit le classeur produit plutot que de refaire la consolidation :
    le chiffre publie et celui de cet etat doivent etre le meme, sinon
    l'un des deux ment.
    """
    chemin = classeur_fusionne()
    print(f"  source des prix : {chemin.name}")
    feuille = load_workbook(chemin_lisible(chemin), read_only=True)["Feuil1"]
    lignes = list(feuille.iter_rows(values_only=True))
    # Le cartouche de couverture occupe les premieres lignes : l'entete
    # est la premiere ligne qui porte « Code article ».
    depart = next(i for i, l in enumerate(lignes)
                  if l and "Code article" in [str(c) for c in l])
    entetes = [str(c) for c in lignes[depart]]
    table = pd.DataFrame(lignes[depart + 1:], columns=entetes)
    table = table.loc[:, ~table.columns.duplicated()]
    table["Code article"] = code_article(table["Code article"])
    table = table[table["Code article"].ne("") & table["Code article"].ne("None")]
    garde = ["Code article", "Origine ligne"] + [c for c in REPRISES
                                                 if c in table.columns]
    return table[garde].drop_duplicates("Code article")


def motifs() -> pd.DataFrame:
    """Pourquoi un article n'a pas de prix, d'apres les balayages.

    Les trois causes ne se valent pas et ne se traitent pas pareil :
    un prix ecarte existe et attend un arbitrage, un article absent du
    tarif demande une verification de reference, un fournisseur sans
    tarif demande un courrier. Les confondre ferait perdre le seul
    classement utile.
    """
    lots = []
    for fichier in sorted(SORTIE_ACHATS.glob("pa_completer_*.xlsx")):
        for feuille, motif in (
                ("Écartés", "prix trouvé mais écarté — remise non portée "
                            "par le tarif, à arbitrer"),
                ("Sans tarif", "fournisseur tarifé, mais l'article n'est "
                               "pas dans son tarif")):
            try:
                lot = pd.read_excel(chemin_lisible(fichier),
                                    sheet_name=feuille, skiprows=3)
            except Exception:
                continue
            if lot.empty or "Code article" not in lot.columns:
                continue
            lots.append(pd.DataFrame({
                "Code article": code_article(lot["Code article"]),
                "Motif": motif,
                "PA écarté (à vérifier)": lot.get("PA écarté (à vérifier)"),
                "Alerte balayage": lot.get("Alerte"),
            }))
    if not lots:
        return pd.DataFrame(columns=["Code article", "Motif"])
    # Un ecart arbitrable prime sur une absence : c'est la piste la plus
    # avancee des deux.
    tout = pd.concat(lots, ignore_index=True)
    tout["_rang"] = tout["Motif"].str.startswith("prix trouvé").map(
        {True: 0, False: 1})
    return (tout.sort_values("_rang").drop_duplicates("Code article")
            .drop(columns="_rang"))


def main() -> None:
    dedans = perimetre_liste.charger().copy()
    print(f"  {perimetre_liste.entete()}")
    dedans["PERIMETRE_VERSION"] = perimetre_liste.manifeste()["version"]

    wms = pd.read_excel(chemin_lisible(EXTRACT), dtype=str,
                        usecols=["Code article", "Nom fabriquant",
                                 "Référence fabricant", "Ref. art. four.",
                                 "Nom fournisseur"])
    wms = wms.dropna(subset=["Code article"]).drop_duplicates("Code article")
    wms["Code article"] = wms["Code article"].str.strip()
    wms = wms.rename(columns={"Nom fournisseur": "Fournisseur (fiche WMS)",
                              "Ref. art. four.": "Réf. article fournisseur"})
    etat = dedans.merge(wms, on="Code article", how="left")

    etat = etat.merge(prix_connus(), on="Code article", how="left")
    etat = etat.merge(motifs(), on="Code article", how="left")

    # L'Ancien PA venait du seul classeur consolide — or un article sans
    # prix 2026 n'y figure pas, et ressortait donc sans prix ACTUEL non
    # plus. C'est precisement l'inverse de ce qu'on veut voir : sur un
    # article qui reste a faire, savoir ce qu'on paie aujourd'hui est la
    # premiere information utile. On le complete depuis le WMS.
    try:
        import pa_completer
        actuels = (pa_completer.pa_actuels()
                   .drop_duplicates("code_article")
                   .set_index("code_article")["pa_wms"])
        if "Ancien PA" not in etat.columns:
            etat["Ancien PA"] = None
        vide = etat["Ancien PA"].isna()
        etat.loc[vide, "Ancien PA"] = etat.loc[vide, "Code article"].map(
            actuels)
        print(f"  Ancien PA complété depuis le WMS sur "
              f"{int(vide.sum() - etat['Ancien PA'].isna().sum())} articles")
    except Exception as err:                       # pragma: no cover
        print(f"  ! PA du WMS indisponible : {type(err).__name__}")

    # LES ARBITRAGES HUMAINS, POSES ICI ET PAS PLUS LOIN (18/09)
    #
    # Ils l'etaient jusqu'ici dans la vue finale seulement. Consequence :
    # AUCUNE autre vue ne les voyait — ni la reconciliation, qui
    # re-signalait a chaque passe des divergences deja tranchees, ni le
    # controle par les commandes, ni les feuilles de decision. Une alarme
    # qui sonne encore apres qu'on a traite le probleme finit par etre
    # ignoree, et c'est ce qu'on veut le moins sur un outil relance a
    # chaque passe.
    #
    # `pa_etat_par_article` etant le socle dont TOUTES les vues derivent,
    # c'est ici que les corrections doivent entrer. Elles le font AVANT le
    # calcul du statut, pour qu'un article corrige cesse d'etre « sans
    # prix », et JAMAIS sur une ligne du responsable des prix.
    try:
        import pa_corrections_17_09 as corrections

        codes = code_article(etat["Code article"])
        est_valide = etat["Origine ligne"].fillna("").str.startswith(
            "Responsable")
        nouveau = pd.to_numeric(etat["Nouveau PA"], errors="coerce")
        ancien = pd.to_numeric(etat["Ancien PA"], errors="coerce")
        poses, gardes = 0, 0

        for entree in corrections.PRIX:
            cible = (codes == str(entree["code_article"])) & ~est_valide
            if not cible.any():
                continue
            etat.loc[cible, "Nouveau PA"] = entree["prix"]
            etat.loc[cible, "Origine Nouveau PA"] = "Arbitrage manuel"
            etat.loc[cible, "Fichier source PA"] = "pa_corrections_17_09.py"
            etat.loc[cible, "Clé de rapprochement"] = "arbitrage manuel"
            poses += int(cible.sum())

        # Les prix écrits directement dans la dernière colonne du livrable,
        # relevés au registre durable. Même règle : jamais sur une ligne
        # déjà validée par le responsable des prix.
        import pa_annotations_livrable

        for code, prix in pa_annotations_livrable.prix_annotes().items():
            cible = (codes == str(code)) & ~est_valide
            if not cible.any():
                continue
            etat.loc[cible, "Nouveau PA"] = float(prix)
            etat.loc[cible, "Origine Nouveau PA"] = (
                "Annotation du livrable")
            etat.loc[cible, "Fichier source PA"] = (
                "pa_annotations_livrable.csv")
            etat.loc[cible, "Clé de rapprochement"] = "annotation livrable"
            poses += int(cible.sum())

        # « GARDE » : le candidat propose etait faux, l'ancien PA reprend
        # sa place. Sans ancien PA il n'y a rien a reprendre — on ne pose
        # pas un prix qu'on n'a pas.
        for entree in corrections.GARDE:
            cible = ((codes == str(entree["code_article"])) & ~est_valide
                     & ancien.notna())
            if not cible.any():
                continue
            etat.loc[cible, "Nouveau PA"] = ancien[cible]
            etat.loc[cible, "Origine Nouveau PA"] = (
                "Arbitrage manuel — ancien PA confirmé")
            etat.loc[cible, "Fichier source PA"] = "pa_corrections_17_09.py"
            gardes += int(cible.sum())

        if poses or gardes:
            # L'ecart se recalcule sur les lignes touchees, sinon il garde
            # la trace du prix d'avant et raconte n'importe quoi.
            n = pd.to_numeric(etat["Nouveau PA"], errors="coerce")
            a = pd.to_numeric(etat["Ancien PA"], errors="coerce")
            calculable = n.notna() & a.notna() & (a != 0)
            etat.loc[calculable, "Écart %"] = ((n - a) / a).round(4)
            print(f"  arbitrages manuels : {poses} prix posés, "
                  f"{gardes} anciens PA confirmés")
    except Exception as err:                           # pragma: no cover
        print(f"  ! arbitrages manuels indisponibles : {type(err).__name__}: "
              f"{err}")

    # Le statut est la seule colonne qui se lit d'un coup d'oeil : trois
    # valeurs, pas davantage, et ce que le responsable des prix a validé y
    # est distingue du reste parce que lui seul a ete relu.
    origine = etat["Origine ligne"].fillna("")
    a_un_prix = pd.to_numeric(etat["Nouveau PA"], errors="coerce").notna()
    etat["Statut PA"] = "3 — SANS PRIX"
    etat.loc[a_un_prix, "Statut PA"] = "2 — PROPOSÉ (à relire)"
    etat.loc[origine.str.startswith("Responsable"), "Statut PA"] = \
        "1 — ACQUIS (validé par le responsable)"

    # Un article sans prix ET sans motif connu n'a jamais ete balaye :
    # son fournisseur n'a envoye aucun tarif. Le dire est plus utile que
    # de laisser la case vide.
    sans_motif = (etat["Statut PA"] == "3 — SANS PRIX") & etat["Motif"].isna()
    sans_fournisseur = (etat["Fournisseur"].fillna("").str.strip() == "")
    etat.loc[sans_motif & ~sans_fournisseur, "Motif"] = \
        "aucun tarif reçu de ce fournisseur"
    etat.loc[sans_motif & sans_fournisseur, "Motif"] = (
        "aucun fournisseur au périmètre — voir le fabricant")
    etat.loc[etat["Statut PA"] != "3 — SANS PRIX", "Motif"] = ""

    # Les fournisseurs qu'on a quittes sortent du chantier : leurs articles ne
    # sont pas « sans prix » en attente d'un tarif, ils sont HORS CHANTIER.
    # La nuance n'est pas cosmetique — un article sans prix appelle un
    # travail, et celui-ci non. Sans elle, ils remontent a chaque passe dans
    # les listes a traiter, et quelqu'un finit par chercher leur tarif.
    from fournisseurs_arretes import est_arrete, motif as motif_arret

    arretes = etat["Fournisseur"].fillna("").map(est_arrete)
    a_marquer = arretes & (etat["Statut PA"] == "3 — SANS PRIX")
    etat.loc[a_marquer, "Motif"] = etat.loc[a_marquer, "Fournisseur"].map(
        motif_arret)
    etat.loc[arretes, "Statut PA"] = "4 — HORS CHANTIER (fournisseur arrêté)"
    if arretes.any():
        print(f"    {int(arretes.sum())} articles hors chantier "
              f"(fournisseur arrêté) — dont "
              f"{int(a_marquer.sum())} qui étaient sans prix")

    # CE QU'ON ÉCARTE N'EST PAS CE QUI RÉSISTE (18/09)
    #
    # « 3 — SANS PRIX » mélangeait deux choses qui n'ont rien à voir : les
    # articles qu'on n'arrive pas à chiffrer, et ceux qu'on a DÉCIDÉ de ne
    # pas traiter — lits, fauteuils roulants, VPH, produits arrêtés. Sur
    # 1 078 « sans prix », 522 étaient dans le second cas. Lire ce chiffre
    # comme un reste-à-faire, c'est se tromper d'un facteur deux, et c'est
    # arrivé le 18/09 dans une note destinée à l'extérieur.
    #
    # La règle n°3 du chantier dit quoi faire : « tout filtre devient une
    # colonne ». La mise à l'écart devient donc une COLONNE DU SOCLE, et
    # non plus une information qu'il faut aller chercher dans un second
    # fichier en recroisant deux sources.
    #
    # Le classement vient de `pa_ecartes.classer()` — la MÊME fonction que
    # celle qui produit les feuilles d'écartement, jamais une copie : deux
    # écritures de la même règle finiraient par se contredire.
    #
    # `Statut PA` n'est pas touché. Les vues existantes qui comptent sur
    # « 3 — SANS PRIX » continuent de fonctionner ; celles qui veulent le
    # reste-à-faire réel filtrent sur « Mise à l'écart » vide.
    try:
        import pa_ecartes

        annot = pa_ecartes.annotations()
        avec = etat.merge(annot, on="Code article", how="left")
        etat["Mise à l'écart"] = pa_ecartes.classer(avec)["motif"].values
        ecartes = etat["Mise à l'écart"].fillna("") != ""
        sans_prix = etat["Statut PA"] == "3 — SANS PRIX"
        print(f"    {int(ecartes.sum())} articles mis à l'écart "
              f"(lit/fauteuil/VPH, fournisseur arrêté, arrêt fabricant)")
        print(f"    reste-à-faire RÉEL : "
              f"{int((sans_prix & ~ecartes).sum())} sans prix au périmètre "
              f"actif, sur {int(sans_prix.sum())} « sans prix » au total")
    except Exception as err:                           # pragma: no cover
        etat["Mise à l'écart"] = ""
        print(f"  ! mise à l'écart indisponible : {type(err).__name__}: {err}")

    colonnes = [
        "PERIMETRE_VERSION", "Réf. interne", "Code article",
        "Code déclinaison", "Désignation", "Type", "Famille",
        "Fournisseur", "Nom fabriquant", "Réf. article fournisseur",
        "Référence fabricant", "Réf. fournisseur",
        "Statut PA", "Mise à l'écart", "Motif",
        "Ancien PA", "Nouveau PA", "Écart %", "Origine ligne",
        "Origine Nouveau PA", "Fichier source PA",
        "Clé de rapprochement", "Alerte", "PA écarté (à vérifier)",
        "Alerte balayage",
        "VENTE_12M", "ACHAT_12M", "PORTE",
    ]
    sortie = etat.reindex(columns=[c for c in colonnes if c in etat.columns])
    for c in ("Ancien PA", "Nouveau PA", "Écart %",
              "PA écarté (à vérifier)"):
        if c in sortie.columns:
            sortie[c] = pd.to_numeric(sortie[c], errors="coerce").round(4)

    total = len(sortie)
    compte = sortie["Statut PA"].value_counts()
    acquis = int(compte.get("1 — ACQUIS (validé par le responsable)", 0))
    propose = int(compte.get("2 — PROPOSÉ (à relire)", 0))
    print(f"\n{total} articles")
    for statut, n in compte.sort_index().items():
        print(f"  {statut:<30} {n:>5}  ({n / total:.1%})")
    print("\nmotifs des articles sans prix :")
    for motif, n in (sortie.loc[sortie["Statut PA"] == "3 — SANS PRIX",
                                "Motif"].value_counts().items()):
        print(f"  {n:>5}  {motif}")

    chemin = mise_en_forme.chemin_ecriture(
        SORTIE_ACHATS / "pa_etat_par_article.xlsx")
    with pd.ExcelWriter(chemin, engine="openpyxl") as writeur:
        sortie.sort_values(["Statut PA", "Fournisseur", "Désignation"]
                           ).to_excel(writeur, sheet_name="Périmètre",
                                      index=False, startrow=3)
        (sortie[sortie["Statut PA"] == "3 — SANS PRIX"]
         .sort_values(["Motif", "Fournisseur"])
         .to_excel(writeur, sheet_name="Sans prix", index=False, startrow=3))
    mise_en_forme.formater(chemin, {
        "Périmètre": {
            "ligne_entete": 4, "figer_colonne": 3,
            "titre": "Prix d'achat 2026 — état article par article",
            "sous_titre": (
                f"{total} articles du périmètre figé — "
                f"{acquis} validés par le responsable ({acquis / total:.1%}), "
                f"{propose} proposés ({propose / total:.1%}), "
                f"{total - acquis - propose} sans prix "
                f"({(total - acquis - propose) / total:.1%}). "
                f"Seul « ACQUIS » a été relu.")},
        "Sans prix": {
            "ligne_entete": 4, "figer_colonne": 3,
            "titre": "Les articles qui n'ont aucun prix 2026",
            "sous_titre": "Triés par motif : c'est le motif qui dit à qui "
                          "revient l'action"},
    })
    print(f"\n  -> {chemin}")


if __name__ == "__main__":
    sys.exit(main())
