# -*- coding: utf-8 -*-
"""Réconciliation entre le classeur du responsable des prix et nos propositions de PA.

POURQUOI
    Deux pipelines lisent les mêmes fichiers tarifs. Celui du responsable des prix
    (./data/) porte 33 fournisseurs dont la feuille, la ligne d'en-tête et l'intitulé
    exact de la colonne prix ont été établis un par un. Le nôtre balaie 154
    fournisseurs. Là où les deux ont un prix pour le même article, le même fournisseur
    et le même palier, ils doivent dire EXACTEMENT la même chose. Toute différence est
    un défaut de lecture d'un des deux côtés, et il faut savoir lequel avant d'ajouter
    un seul prix de plus.

CE QUE LE SCRIPT NE FAIT PAS
    Il ne corrige rien et n'écrit dans aucun fichier de travail. Il produit un classeur
    de constat. Le responsable des prix fait foi : on ne touche pas à ses valeurs, on
    documente l'écart.

LECTURE DU RÉSULTAT
    Le classeur sortie/3-achats/reconciliation_responsable.xlsx porte six onglets :

      Synthese        les compteurs, à lire en premier
      Par fournisseur rapport médian nous / responsable, fournisseur par fournisseur.
                      C'est l'onglet qui trouve les erreurs de colonne : un rapport
                      médian stable sur des dizaines de lignes n'est pas trente
                      désaccords, c'est une colonne différente ou une unité différente.
      Divergences     les deux ont un prix et ils diffèrent, segmenté par origine
      Seul responsable  le responsable a un prix, nous non. Pourquoi notre balayage
                      l'a-t-il manqué ?
      Seul nous       notre apport net, ce qui est réellement à faire valider
      Sans prix       ni l'un ni l'autre, c'est le reste du chantier

USAGE
    python pa_reconcilier_responsable.py [--perimetre chemin.csv]
                                  [--responsable chemin.xlsx]
                                  [--notre chemin.xlsx] [--sortie chemin.xlsx]

    Les classeurs peuvent rester ouverts dans Excel, ils sont lus depuis une copie.
"""
import argparse
import os
import re
import shutil
import sys
import tempfile

import numpy as np
import pandas as pd

# Tolérance d'égalité : reprise de la convention de verif_ecarts.py.
# Deux prix issus de la même cellule doivent coller au centime, pas à 2 %.
def _egaux(a, b):
    return (a - b).abs() <= np.maximum(0.005, 0.0005 * b.abs())


# Rapport médian par fournisseur au-delà duquel on parle de colonne mal choisie
# plutôt que de désaccord ligne à ligne.
MIN_LIGNES_RATIO = 5
TOLERANCE_RATIO = 0.005


# --------------------------------------------------------------------------- entrées/sorties
def chemin_lisible(chemin):
    """Copie temporaire : un classeur ouvert dans Excel reste lisible."""
    if not os.path.exists(chemin):
        raise FileNotFoundError(chemin)
    tmp = os.path.join(tempfile.gettempdir(),
                       "recon_" + os.path.basename(chemin))
    shutil.copy2(chemin, tmp)
    return tmp


def lire(chemin, onglet, temoin="Code article", max_lignes=8):
    """Lit un onglet en trouvant sa ligne d'en-tête.

    Les classeurs de ce projet portent un cartouche de titre sur les trois
    premières lignes et leur en-tête en ligne 4 ; ceux du responsable des prix
    commencent directement par l'en-tête. On cherche donc la première ligne qui contient
    la colonne témoin plutôt que de supposer l'une ou l'autre convention.
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
    """Résout un nom de colonne parmi plusieurs orthographes possibles.

    Échoue bruyamment en listant les colonnes réellement présentes : c'est la seule
    façon de s'apercevoir qu'un onglet a été renommé plutôt que de produire un
    résultat vide sans erreur.
    """
    normal = {re.sub(r"\s+", " ", str(c)).strip().lower(): c for c in df.columns}
    for cand in candidats:
        k = re.sub(r"\s+", " ", cand).strip().lower()
        if k in normal:
            return normal[k]
    if not obligatoire:
        return None
    raise KeyError("%s : aucune colonne parmi %s. Présentes : %s"
                   % (ou or "fichier", list(candidats), list(df.columns)))


def code(s):
    """Normalise un code article. pandas transforme 32818 en « 32818.0 » dès qu'une
    valeur manque dans la colonne, et la jointure tombe alors à zéro sans lever
    la moindre erreur."""
    return (s.astype(str).str.strip()
             .str.replace(r"\.0$", "", regex=True)
             .str.lstrip("/")
             .replace({"nan": np.nan, "None": np.nan, "": np.nan}))


def frs(s):
    """Normalise un libellé fournisseur pour la comparaison."""
    return (s.astype(str).str.upper().str.strip()
             .str.replace(r"[^A-Z0-9]+", " ", regex=True)
             .str.replace(r"\s+", " ", regex=True)
             .replace({"NAN": np.nan, "": np.nan}))


def palier(s):
    """Palier numérique. Un palier absent n'est PAS un palier 1 : la distinction est
    conservée, c'est elle qui explique une partie des écarts."""
    return pd.to_numeric(s, errors="coerce")


# --------------------------------------------------------------------------- chargements
def charger_perimetre(chemin):
    df = pd.read_csv(chemin, sep=None, engine="python", dtype=str)
    c_art = col(df, "Code article", "code_article", ou="périmètre")
    c_dec = col(df, "Code déclinaison", "Code declinaison", "code_declinaison",
                obligatoire=False)
    out = pd.DataFrame({"art": code(df[c_art])})
    out["dec"] = code(df[c_dec]) if c_dec else ""
    out["dec"] = out["dec"].fillna("")
    out = out.dropna(subset=["art"]).drop_duplicates()
    print("Périmètre        : %d articles" % out["art"].nunique())
    return out


def charger_responsable(chemin, onglet="Base articles"):
    df = lire(chemin, onglet)
    ou = "classeur du responsable / " + onglet
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
    print("Responsable      : %d lignes, %d avec un Nouveau PA"
          % (len(out), out["pa_resp"].notna().sum()))
    return out


def charger_notre(chemin, onglet=0):
    df = lire(chemin, onglet if onglet else 0)
    ou = "nos propositions"
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
    print("Nos propositions : %d lignes, %d avec un prix"
          % (len(out), out["pa_nous"].notna().sum()))
    return out


# --------------------------------------------------------------------------- réconciliation
def palier_unitaire_du_responsable(responsable):
    """Réduit le responsable à UNE ligne par article/déclinaison/fournisseur.

    À utiliser seulement quand notre côté ne porte pas de palier. Nos propositions
    sont des prix UNITAIRES, et un prix unitaire ne se compare qu'à un prix
    unitaire. On retient donc, dans cet ordre :

        1. la ligne de palier 1 — le prix à l'unité, celui qui est comparable ;
        2. à défaut, la ligne SANS palier, qui vaut le plus souvent pour l'unité.

    Tout le reste est marqué NON COMPARABLE. Chez les Fournisseurs P, AP ou CD,
    le plus petit palier du responsable est 5, 6, parfois 28 : son prix y est
    dégressif. Le confronter à notre prix unitaire produirait un écart qui ne dit
    rien du tarif — un FAUX ÉCART, dans l'autre sens que celui du 13/08. Mieux vaut
    une ligne classée « non comparable » qu'une divergence inventée.
    """
    grp = ["art", "dec", "frs"]
    # Ordre de préférence : palier 1, puis palier absent, puis le reste par
    # valeur croissante — seul un tri explicite garantit lequel survit.
    rang = np.where(responsable["palier"] == 1, 0,
                    np.where(responsable["palier"].isna(), 1, 2))
    j = responsable.assign(_rang=rang).sort_values(["_rang", "palier"],
                                                   na_position="last")
    # Combien de paliers cet article porte-t-il chez le responsable ? Au-delà de
    # un, la comparaison ne vaut que pour celui qu'on retient, et il faut le dire.
    j["Paliers chez le responsable"] = j.groupby(
        grp, dropna=False)["pa_resp"].transform("size")
    j = j.drop_duplicates(grp, keep="first")
    # Un prix unitaire existe, ou il n'existe pas : c'est cette colonne qui
    # décidera ensuite si l'on a le droit de parler d'écart.
    j["comparable_unitaire"] = j["_rang"] < 2
    return j.drop(columns="_rang").rename(
        columns={"palier": "Palier responsable retenu"})


def reconcilier(perim, responsable, notre):
    # Restriction au périmètre figé. On le LIT, on ne le recalcule pas.
    responsable = responsable.merge(perim, on=["art", "dec"], how="inner")
    notre = notre.merge(perim, on=["art", "dec"], how="inner")

    # La clé R4 du responsable — article, déclinaison, fournisseur, palier — suppose
    # que les DEUX côtés portent un palier. Notre pa_etat_par_article.xlsx est une vue
    # PAR ARTICLE : sa colonne palier est vide de bout en bout. Joindre dessus ne
    # rapprochait que les 61 lignes où le responsable était lui aussi à vide, et le
    # script concluait « aucune divergence » après n'avoir comparé presque rien — un
    # feu vert non mérité, plus dangereux qu'une divergence signalée.
    #
    # On choisit donc la clé en fonction de ce que les données portent vraiment,
    # et on le dit à l'écran plutôt que de le supposer.
    au_palier = "palier" in notre.columns and notre["palier"].notna().any()
    if au_palier:
        cle = ["art", "dec", "frs", "palier"]
        print("Clé de comparaison : article + déclinaison + fournisseur + palier")
    else:
        cle = ["art", "dec", "frs"]
        responsable = palier_unitaire_du_responsable(responsable)
        print("Clé de comparaison : article + déclinaison + fournisseur")
        print("  nos prix sont UNITAIRES et sans palier : on les confronte au "
              "palier le plus bas du responsable")
        multi = int((responsable["Paliers chez le responsable"] > 1).sum())
        if multi:
            print("  %d articles portent plusieurs paliers chez le responsable "
                  "— la comparaison ne vaut que pour celui retenu, dit en colonne"
                  % multi)

    m = responsable.merge(notre, on=cle, how="outer", indicator=True,
                          suffixes=("_j", "_n"))

    a_resp = m["pa_resp"].notna()
    a_nous = m["pa_nous"].notna()

    # En mode unitaire, les deux prix ne sont opposables que si celui du responsable
    # est bien un prix à l'unité. Sinon l'écart n'est pas un désaccord de lecture,
    # c'est une dégressivité — et la nommer autrement ferait courir après un
    # défaut qui n'existe pas.
    if "comparable_unitaire" in m.columns:
        # .astype(bool) n'est pas cosmétique : la jointure externe introduit des
        # NaN, la colonne retombe en dtype object, et « ~True » y vaut -2 —
        # un entier non nul, donc vrai. Sans ce cast, toute ligne devient « non
        # comparable » sans qu'aucune erreur ne soit levée.
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

    # Segmentation des divergences par origine du prix du responsable. Une divergence
    # contre un tarif fournisseur est un défaut de lecture. Une divergence contre le
    # fichier interne est un arbitrage commercial, et ce n'est pas le même travail.
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
    """Rapport médian nous / responsable par fournisseur, et par colonne de prix lue.

    Un rapport médian stable et différent de 1 sur au moins quelques lignes ne se
    négocie pas ligne à ligne : c'est une colonne ou une unité, et ça se règle en un
    seul arbitrage pour tout le fournisseur.
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


# --------------------------------------------------------------------------- sortie
# « Palier responsable retenu » et « Paliers chez le responsable » n'existent que
# dans le mode de comparaison unitaire. Ils disent sur QUELLE ligne du responsable
# l'écart a été calculé : sans eux, un écart contre un article à six paliers serait
# illisible.
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
    if os.path.exists(chemin):                      # le classeur peut être ouvert
        base, ext = os.path.splitext(chemin)
        try:
            os.remove(chemin)
        except PermissionError:
            chemin = "%s %s%s" % (base, pd.Timestamp.now().strftime("%Y-%m-%d %H%M"), ext)
    with pd.ExcelWriter(chemin, engine="openpyxl") as w:
        # sheet_name est un argument NOMMÉ depuis pandas 2 : le passer en
        # position lève un TypeError après avoir déjà ouvert le classeur.
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
        # Ni un accord ni un désaccord : le responsable n'a que des paliers
        # dégressifs sur ces articles. À reprendre quand nos propositions
        # porteront un palier.
        _vue(m, "PALIER NON UNITAIRE").to_excel(
            w, sheet_name="Palier non unitaire", index=False)
        _vue(m, "SANS PRIX").to_excel(w, sheet_name="Sans prix", index=False)
    return chemin


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--perimetre", default=os.path.join("perimetre", "perimetre_v1_2.csv"))
    # Le classeur de reference est la BASE du responsable des prix, celle de
    # 30 197 lignes qui porte l'onglet « Base articles ». Deux pieges evites ici :
    #   - le dossier par defaut n'existait pas sur le poste de travail, et
    #     l'ancien defaut ne pouvait qu'echouer ;
    #   - « prix_valides.xlsx » n'a PAS d'onglet « Base articles » (le sien
    #     s'appelle Feuil1), donc chemin et onglet par defaut se
    #     contredisaient.
    # On vise la base SOURCE, jamais une copie « + propositions » : reconcilier
    # contre un fichier ou nos propres prix ont deja ete verses reviendrait a
    # se comparer a soi-meme et a n'y trouver aucun ecart.
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
            print("FOURNISSEURS À REPRENDRE")
            print(suspects.to_string(index=False))
        else:
            print("Aucun fournisseur en divergence.")
    chemin = ecrire(m, perim, a.sortie)
    print()
    print("Écrit : %s" % chemin)

    # Un défaut de lecture n'est pas une nuance : tant qu'il en reste, aucune
    # proposition supplémentaire ne doit partir chez le responsable des prix.
    lecture = int((m["Nature"].str.startswith("LECTURE")).sum())
    if lecture:
        print("!! %d divergence(s) de LECTURE : à trancher avant d'étendre le chantier"
              % lecture)
        sys.exit(1)


if __name__ == "__main__":
    main()
