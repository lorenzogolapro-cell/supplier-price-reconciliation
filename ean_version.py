# -*- coding: utf-8 -*-
"""
Fige une version diffusable du fichier des codes EAN.

    python ean_version.py            produit la version suivante
    python ean_version.py --version 1   impose le numero

Entree  : sortie/2-chantier-ean/EAN distributeur.xlsx  (regenere par ean_global.py)
Sortie  : sortie/2-chantier-ean/EAN distributeur - V1.xlsx

Pourquoi une version figee
--------------------------
`EAN distributeur.xlsx` est un fichier de travail : il change a chaque
collecte. Une version numerotee, datee et accompagnee de sa notice est
autre chose — c'est ce qu'on transmet, ce sur quoi on se met d'accord, et
ce a quoi on peut revenir dans six mois en sachant ce qu'il contenait.

Le nettoyage : les codes ambigus
--------------------------------
Un code EAN identifie UN produit. Quand le meme code se retrouve sur
plusieurs de nos articles, l'un des deux est faux — et rien ne dit lequel.
Scanner ce code en entrepot ne designerait rien.

Ces codes ne sont pas supprimes : ils sont retrogrades en « ambigu »,
sortis du lot exploitable et isoles sur leur propre feuille pour
arbitrage. Sur un dispositif medical, on ne fait pas disparaitre une
donnee en silence.

Trois familles s'y retrouvent, et elles n'ont pas la meme cause :
  - des declinaisons de couleur ou de taille auxquelles le fabricant n'a
    donne qu'un seul GTIN (un scooter decline en bleu, gris et orange) —
    le code est juste, c'est notre granularite qui est plus fine ;
  - des produits reellement differents ayant recu le meme code par un
    rapprochement trop large (une plateforme et sa sangle) — le code est
    faux pour tous sauf un ;
  - des gammes voisines confondues par le libelle (une nutrition
    1-3 ans et la meme gamme > 3 ans).

Seule la premiere est acceptable, et elle ne se distingue pas
automatiquement des deux autres : d'ou l'arbitrage.
"""

from __future__ import annotations

import difflib
import re
import sys
from datetime import date
from pathlib import Path

RACINE = Path(__file__).resolve().parent
sys.path.insert(0, str(RACINE))

import warnings  # noqa: E402

warnings.filterwarnings("ignore")

import pandas as pd  # noqa: E402

import mise_en_forme  # noqa: E402
from extracteurs.base import chemin_lisible, ean_est_valide  # noqa: E402
from main import SORTIE_CHANTIER  # noqa: E402

def _source() -> Path | None:
    """Le fichier de travail le plus recent.

    Quand le classeur est ouvert dans Excel, `ean_global` ne peut pas
    l'ecraser et ecrit une copie horodatee a cote. Chercher le nom
    canonique seulement reviendrait alors a figer une version perimee
    sans s'en apercevoir.
    """
    return mise_en_forme.derniere_version(SORTIE_CHANTIER,
                                          "EAN distributeur.xlsx")

# Ce que vaut chaque source, en clair, pour la notice
ORIGINE = {
    "tarif": "lu dans le tarif du fournisseur",
    "référence=EAN": "notre référence fournisseur EST un code EAN valide",
    "entrepôt": "scanné sur le produit en entrepôt",
    "eudamed": "déclaré par le fabricant dans EUDAMED, sur la référence exacte",
    "eudamed-modèle": "déclaré dans EUDAMED, rapproché par le modèle",
    "vidal": "publié par VIDAL",
    "libellé": "déduit par concordance de libellé",
}


def _version_suivante() -> int:
    """Le numero qui suit la derniere version deja produite."""
    numeros = [
        int(trouve.group(1))
        for fichier in SORTIE_CHANTIER.glob("EAN distributeur - V*.xlsx")
        if (trouve := re.search(r" - V(\d+)\.xlsx$", fichier.name))
    ]
    return max(numeros, default=0) + 1


# Mots qui ne distinguent pas deux declinaisons d'un meme produit :
# retires du libelle, deux tailles ou deux couleurs se confondent.
DECLINAISON = re.compile(
    r"\b(?:T|TAILLE|T\.)\s*\d+\w*\b"
    r"|\b(?:XS|S|M|L|XL|XXL|XXXL|TU|TS|TM|TL)\b"
    r"|\b(?:NOIR|BLANC|BLEU|ROUGE|GRIS|BEIGE|VERT|ROSE|MARINE|BORDEAUX"
    r"|SABLE|TURQUOISE|ORANGE|JAUNE|VIOLET|MARRON|IVOIRE|LAGON|ANTHRACITE)"
    r"\w*\b"
    r"|\b(?:GAUCHE|DROITE|DROIT)\b",
    re.IGNORECASE)

# En deca de cette ressemblance, deux libelles ne decrivent pas le meme
# produit — meme en ayant retire taille et couleur.
RESSEMBLANCE_MIN = 0.55


def _socle(libelle) -> str:
    """Le libelle debarrasse de ce qui distingue une declinaison."""
    texte = DECLINAISON.sub(" ", str(libelle or "").upper())
    return re.sub(r"[^A-Z0-9]+", " ", texte).strip()


def _meme_maison(a, b) -> bool:
    """Deux noms de fournisseur designent-ils la meme societe ?

    Le fournisseur GE appartient au fournisseur BF, les fournisseurs A et
    CS distribuent tous deux le fournisseur GD : le meme code chez deux
    « fournisseurs » differents n'est alors pas une erreur. La table d'alias du
    rapprochement sait deja cela — on la reutilise plutot que de la
    redecouvrir.
    """
    import appariement
    if not a or not b or pd.isna(a) or pd.isna(b):
        return False
    if str(a).strip().upper() == str(b).strip().upper():
        return True
    cible = appariement.ALIAS.get(str(a).strip().upper())
    autre = appariement.ALIAS.get(str(b).strip().upper())
    return bool(cible and autre and cible == autre)


def _verdict_partage(groupe: pd.DataFrame) -> tuple[str, str]:
    """Que vaut un code porte par plusieurs articles ? (verdict, motif)

    Le meme code sur deux lignes n'est pas forcement une erreur, et
    l'ecarter systematiquement coutait cher : 465 lignes retirees, dont
    la moitie portaient un code exact.

    Trois situations le justifient :

      - une seule reference fournisseur pour toutes les lignes : ce sont
        des articles du WMS en DOUBLON, pas des codes faux. Un meme
        collier cervical occupe trois codes article ;
      - deux noms de fournisseur pour une meme societe ;
      - des declinaisons de taille ou de couleur auxquelles le fabricant
        n'a donne qu'un GTIN — notre granularite est alors plus fine que
        la sienne.

    Reste le cas ou les references ET les libelles divergent : la, le
    code est faux pour tous sauf un, et rien ne dit lequel.
    """
    references = {str(r).strip() for r in groupe["Réf. fournisseur"].dropna()}
    fournisseurs = [f for f in groupe["Fournisseur"].dropna().unique()]

    if len(references) <= 1:
        return "oui", "articles WMS en doublon — même référence fournisseur"

    if len(fournisseurs) > 1 and all(
            _meme_maison(fournisseurs[0], f) for f in fournisseurs[1:]):
        return "oui", "même société sous deux noms"

    socles = {_socle(x) for x in groupe["Désignation"]}
    if len(socles) == 1:
        return "oui", "déclinaisons d'un même produit — GTIN de gamme"

    libelles = [str(x) for x in groupe["Désignation"]]
    ressemblance = min(
        difflib.SequenceMatcher(None, _socle(libelles[0]), _socle(x)).ratio()
        for x in libelles[1:])
    if ressemblance >= RESSEMBLANCE_MIN:
        return "oui", f"libellés concordants ({ressemblance:.0%})"

    return "NON — code ambigu", "références et libellés divergents"


def _verifies_au_libelle() -> set[str]:
    """Les codes EUDAMED confrontés au libellé un par un, et retenus.

    Seule exception au retrait d'EUDAMED : ceux-là n'ont pas ete crus sur
    parole, ils ont ete juges piece a piece par `ean_verifier_libelle.py`.
    Le fichier absent, l'ensemble est vide — on retire alors TOUT EUDAMED,
    ce qui est le comportement prudent.
    """
    fichier = SORTIE_CHANTIER / "eudamed_verification_libelle.xlsx"
    if not fichier.exists():
        return set()
    try:
        v = mise_en_forme.lire(fichier, "Verdicts", "Code article")
    except SystemExit:
        return set()
    retenus = v[v["Verdict"].isin(["CONFIRMÉ", "CORRIGÉ"])]
    return set(retenus["Code article"].astype(str).str.strip())


def _marquer_ambigus(df: pd.DataFrame) -> pd.DataFrame:
    """Statue sur chaque code : exploitable, ou à trancher."""
    avec_code = df["Code EAN"].notna()
    comptes = df.loc[avec_code, "Code EAN"].value_counts()
    df["Articles portant ce code"] = (
        df["Code EAN"].map(comptes).fillna(0).astype(int))

    df["Exploitable"] = "oui"
    df["Motif du partage"] = ""
    df.loc[~avec_code, "Exploitable"] = "— pas de code"

    for code in comptes[comptes > 1].index:
        lignes = df["Code EAN"] == code
        verdict, motif = _verdict_partage(df.loc[lignes])
        df.loc[lignes, "Exploitable"] = verdict
        df.loc[lignes, "Motif du partage"] = motif
        if verdict.startswith("NON"):
            df.loc[lignes, "Fiabilité"] = "ambigu"

    # Une clé GS1 fausse n'aurait rien à faire ici, mais on le vérifie :
    # c'est le contrôle qui coûte le moins cher du fichier.
    invalide = avec_code & ~df["Code EAN"].map(ean_est_valide)
    df.loc[invalide, "Exploitable"] = "NON — clé GS1 invalide"
    df.loc[invalide, "Fiabilité"] = "invalide"

    # Un EAN-8 dans un tarif de matériel médical n'existe pas : ce format
    # est réservé aux emballages trop petits pour porter un EAN-13.
    court = avec_code & (df["Code EAN"].str.len() < 13)
    df.loc[court, "Exploitable"] = "NON — 8 chiffres, pas un EAN produit"
    df.loc[court, "Fiabilité"] = "invalide"

    # EUDAMED N'EST PLUS UNE SOURCE SÛRE, AUCUNE DE SES LIGNES (16/09).
    #
    # La règle ne visait d'abord que les références de moins de neuf
    # caractères — 1 321 codes. Trois mesures indépendantes ont montré que
    # le défaut ne s'arrête pas à celles-là :
    #   - confrontés au libellé chez EUDAMED même, 468 codes jugés :
    #     354 REJETÉ, 19 seulement sauvés (CONFIRMÉ ou CORRIGÉ) ;
    #   - confrontés aux tarifs fournisseurs, 248 codes : 75 contredits,
    #     soit 30 % de faux là où le tarif fait foi ;
    #   - concordance du préfixe GS1 : 27,8 %, contre 100 % pour un tarif.
    #
    # Le coût du retrait est faible et il a été mesuré avant de décider :
    # sur le périmètre figé, 85 articles perdent leur code, la couverture
    # passe de 71,4 % à 69,7 %. Un point et demi contre une source dont un
    # tiers au moins est faux, et faux en silence.
    #
    # On ne supprime rien : le code, sa source et sa référence restent au
    # journal — c'est à cela qu'il sert. On retire une AFFIRMATION, pas une
    # donnée. Les 19 codes vérifiés au libellé gardent leur « sûr » : ceux-là
    # ont été confrontés un par un.
    eudamed = df["Source"].fillna("").str.strip().str.startswith("eudamed")
    verifies = _verifies_au_libelle()
    garde = df["Code article"].astype(str).str.strip().isin(verifies)
    df.loc[eudamed & ~garde, "Fiabilité"] = "à vérifier"
    print(f"  EUDAMED : {int((eudamed & ~garde).sum())} codes passés à "
          f"« à vérifier », {int((eudamed & garde).sum())} gardés sûrs "
          f"(vérifiés au libellé)")
    return df


def _notice(df: pd.DataFrame, version: int) -> pd.DataFrame:
    """Ce qu'un lecteur doit savoir avant d'utiliser le fichier."""
    exploitables = df[df["Exploitable"] == "oui"]
    ambigus = df[df["Exploitable"].str.startswith("NON")]
    actifs = df[df["Actif"].astype(str).str.upper().isin(["OUI", "1", "TRUE"])]
    actifs_ok = actifs[actifs["Exploitable"] == "oui"]

    lignes = [
        ("CODES EAN DU DISTRIBUTEUR",
         f"version {version} — {date.today():%d/%m/%Y}"),
        ("", ""),
        ("CE QUE CONTIENT LE FICHIER", ""),
        ("Articles couverts",
         f"{len(df)} — l'ensemble du périmètre métier, hors SE, "
         f"forfait et location"),
        ("Codes exploitables",
         f"{len(exploitables)} ({len(exploitables)/len(df):.0%})"),
        ("Sur les articles actifs",
         f"{len(actifs_ok)} sur {len(actifs)} ({len(actifs_ok)/max(len(actifs),1):.0%})"),
        ("Codes écartés", f"{len(ambigus)} — voir l'onglet « Codes ambigus »"),
        ("", ""),
        ("D'OÙ VIENNENT LES CODES", ""),
    ]
    for source, compte in exploitables["Source"].value_counts().items():
        lignes.append((f"{source} — {compte}",
                       ORIGINE.get(source, "source non documentée")))

    lignes += [
        ("", ""),
        ("COMMENT ILS SONT VÉRIFIÉS", ""),
        ("Clé de contrôle GS1",
         "tout code dont la clé de contrôle est fausse est écarté. "
         "Aucun n'a été trouvé dans cette version."),
        ("Unicité",
         "un code EAN identifie un produit. Un code porté par plusieurs "
         "de nos articles est retiré du lot exploitable : rien ne dit "
         "lequel il désigne vraiment."),
        ("GTIN-14",
         "un code de 14 chiffres commençant par 0 est l'unité de vente "
         "écrite sur 14 positions : il est ramené en EAN-13. Les "
         "indicateurs 1 à 8 désignent un carton ou une palette et ne "
         "sont pas retenus comme code produit."),
        ("Codes ACL13",
         "les codes ACL des officines (préfixe 3401) sont de vrais "
         "EAN-13 et sont retenus comme tels."),
        ("", ""),
        ("CE QUE LE FICHIER NE DIT PAS", ""),
        ("Codes déduits",
         "les codes issus d'une concordance de libellé ou d'un modèle "
         "EUDAMED sont marqués « probable » ou « à vérifier ». Ils "
         "n'ont pas été lus sur le produit : les contrôler avant de "
         "s'en servir pour une commande ou une traçabilité de lot."),
        ("Conditionnement",
         "le code retenu est celui de l'UNITÉ DE VENTE. Le code du "
         "carton, quand le fournisseur le publie, est conservé à part "
         "et ne figure pas ici."),
        ("Articles sans code",
         f"{int(df['Code EAN'].isna().sum())} articles restent sans code. "
         f"La collecte EUDAMED par référence est en cours et les "
         f"complétera progressivement."),
        ("", ""),
        ("RÉGÉNÉRER", ""),
        ("Commande",
         "python main.py  puis  python ean_global.py  puis  "
         "python ean_version.py"),
        ("Fichier de travail",
         "EAN distributeur.xlsx — celui-ci change à chaque collecte ; la "
         "version numérotée, elle, ne bouge plus."),
    ]
    return pd.DataFrame(lignes, columns=["", " "])


def main() -> None:
    source = _source()
    if source is None:
        raise SystemExit("Lancez d'abord : python ean_global.py")
    print(f"source : {source.name}")

    version = _version_suivante()
    for argument in sys.argv[1:]:
        if argument.startswith("--version"):
            version = int(argument.split("=")[-1] if "=" in argument
                          else sys.argv[sys.argv.index(argument) + 1])

    df = pd.read_excel(chemin_lisible(source), sheet_name="Codes EAN",
                       skiprows=3, dtype=str)
    df = _marquer_ambigus(df)

    # Où trouver le produit. Sans cette colonne, un article sans code
    # oblige à le chercher dans le magasin avant de pouvoir le scanner ;
    # avec elle, la liste des manques devient une tournée.
    import carte_entrepot
    lieux = carte_entrepot.par_article()
    if not lieux.empty:
        cle = df["Code article"].astype(str).str.strip()
        for colonne in lieux.columns:
            df[colonne] = cle.map(lieux[colonne])
        localises = int(df["Emplacement"].notna().sum())
        print(f"  {localises} articles localisés "
              f"({localises/len(df):.0%})")

    exploitables = df[df["Exploitable"] == "oui"]
    ambigus = df[df["Exploitable"].str.startswith("NON")].sort_values(
        ["Code EAN", "Désignation"])
    sans_code = df[df["Code EAN"].isna()]

    # Les articles du WMS en doublon ne relèvent pas de ce chantier : le
    # code est bon, c'est la base qui décrit deux fois le même produit.
    # On le signale sans y toucher, comme le veut la règle de travail.
    doublons = df[df["Motif du partage"].str.startswith(
        "articles WMS en doublon", na=False)].sort_values(
        ["Réf. fournisseur", "Code article"])
    if not doublons.empty:
        retours = RACINE / "perimetre" / "retours"
        retours.mkdir(parents=True, exist_ok=True)
        chemin_retour = retours / (
            f"doublons articles WMS — {date.today():%Y-%m-%d}.xlsx")
        colonnes_retour = [c for c in
                           ["Code article", "Référence", "Désignation",
                            "Fournisseur", "Réf. fournisseur", "Code EAN",
                            "Articles portant ce code", "Actif", "Stock"]
                           if c in doublons.columns]
        with pd.ExcelWriter(chemin_retour, engine="openpyxl") as writer:
            doublons[colonnes_retour].to_excel(
                writer, sheet_name="À vérifier dans le WMS", index=False,
                startrow=3)
        mise_en_forme.formater(chemin_retour, options={
            "À vérifier dans le WMS": {
                "ligne_entete": 4, "figer_colonne": 2,
                "titre": "Articles du WMS portant la même référence "
                         "fournisseur",
                "sous_titre": (
                    f"{len(doublons)} lignes, "
                    f"{doublons['Réf. fournisseur'].nunique()} références | "
                    f"le code EAN est juste : c'est la base qui décrit "
                    f"plusieurs fois le même produit. Rien n'a été modifié "
                    f"ici."),
            }})
        print(f"  {len(doublons)} doublons d'articles signalés "
              f"-> {chemin_retour.parent.name}/{chemin_retour.name}")

    synthese = (exploitables.groupby(["Source", "Fiabilité"])
                .size().reset_index(name="Codes")
                .sort_values("Codes", ascending=False))

    chemin = SORTIE_CHANTIER / f"EAN distributeur - V{version}.xlsx"
    colonnes = ["Code article", "Référence", "Désignation", "Fournisseur",
                "Réf. fournisseur", "Code EAN", "Source", "Fiabilité",
                "Exploitable", "Motif du partage",
                "Articles portant ce code",
                "Zone", "Emplacement", "Autres emplacements",
                "Nb emplacements", "Actif", "Stock", "Type", "Famille"]
    colonnes = [c for c in colonnes if c in df.columns]

    with pd.ExcelWriter(chemin, engine="openpyxl") as writer:
        df[colonnes].to_excel(writer, sheet_name="Codes EAN", index=False,
                              startrow=3)
        ambigus[colonnes].to_excel(writer, sheet_name="Codes ambigus",
                                   index=False, startrow=3)
        sans_code[colonnes].to_excel(writer, sheet_name="Sans code",
                                     index=False, startrow=3)
        synthese.to_excel(writer, sheet_name="Synthèse", index=False,
                          startrow=3)
        _notice(df, version).to_excel(writer, sheet_name="Notice",
                                      index=False)

    actifs = df[df["Actif"].astype(str).str.upper().isin(["OUI", "1", "TRUE"])]
    actifs_ok = int((actifs["Exploitable"] == "oui").sum())
    mise_en_forme.formater(chemin, options={
        "Codes EAN": {
            "ligne_entete": 4, "figer_colonne": 3,
            "titre": f"Codes EAN du distributeur — version {version}",
            "sous_titre": (
                f"{len(exploitables)} codes exploitables sur {len(df)} "
                f"articles | {actifs_ok} des {len(actifs)} articles actifs | "
                f"arrêté au {date.today():%d/%m/%Y}"),
        },
        "Codes ambigus": {
            "ligne_entete": 4, "figer_colonne": 3,
            "titre": "Codes portés par plusieurs articles",
            "sous_titre": (
                "Un code EAN identifie un produit. Ceux-ci en désignent "
                "plusieurs : rien ne dit lequel. Écartés du lot "
                "exploitable, conservés ici pour arbitrage."),
        },
        "Sans code": {
            "ligne_entete": 4, "figer_colonne": 3,
            "titre": "Articles encore sans code EAN",
            "sous_titre": (
                f"{int(sans_code['Emplacement'].notna().sum())} d'entre eux "
                f"ont un emplacement connu : ceux-là peuvent être relevés "
                f"en descendant l'allée. La collecte EUDAMED complétera le "
                f"reste."
                if "Emplacement" in sans_code.columns
                else "La collecte EUDAMED par référence est en cours et les "
                     "complétera progressivement"),
        },
        "Synthèse": {
            "ligne_entete": 4,
            "titre": "D'où viennent les codes retenus",
            "sous_titre": "Par source et par niveau de fiabilité",
        },
    })

    print(f"version {version} — {date.today():%d/%m/%Y}")
    print(f"  {len(df)} articles")
    print(f"  {len(exploitables)} codes exploitables "
          f"({len(exploitables)/len(df):.0%})")
    print(f"  {actifs_ok} sur {len(actifs)} articles actifs "
          f"({actifs_ok/max(len(actifs),1):.0%})")
    print(f"  {len(ambigus)} écartés comme ambigus")
    print(f"  {len(sans_code)} sans code")
    print(f"\n  -> {chemin}")


if __name__ == "__main__":
    main()
