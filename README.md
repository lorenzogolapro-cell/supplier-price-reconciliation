# Pipeline de réconciliation de tarifs fournisseurs

Pipeline Python qui reconstitue, à partir de tarifs fournisseurs hétérogènes, le code
EAN et le prix d'achat de chaque référence d'un distributeur de dispositifs médicaux,
puis les rapproche de son système de gestion d'entrepôt.

---

> ### Avertissement
>
> **Extrait représentatif d'un système plus large.** Ce dépôt ne s'exécute pas en
> l'état : il ne contient aucune donnée confidentielle et certains modules
> internes ont été volontairement exclus. Le code est fourni pour lecture, pour
> illustrer la démarche et l'architecture. Tous les noms (client, fournisseurs,
> personnes) et tous les chiffres commerciaux sont anonymisés ou remplacés par des
> exemples.

---

## Le problème

Un distributeur reçoit chaque année les tarifs de près d'une centaine de
fournisseurs, dans autant de formats : classeurs Excel aux colonnes toutes
différentes, PDF natifs, PDF scannés, documents de conditions commerciales qui ne
sont pas des tarifs. En face, son référentiel produit interne connaît mal les codes
EAN et porte des prix d'achat périmés. Rapprocher les deux à la main est hors
d'atteinte à cette échelle, et le faire naïvement est pire que de ne rien faire :
une erreur de lecture ne produit pas un plantage, elle produit un prix faux que
personne ne voit passer.

## Architecture

Le pipeline se lit en cinq temps. Les noms ci-dessous sont ceux des fichiers de ce
dépôt.

**1. Extraction.** Un module par fournisseur, tous tenus au même contrat :
`extract(path) -> DataFrame` respectant le schéma commun.

- `extracteurs/base.py` : le socle. Schéma de sortie, validation de la clé de
  contrôle EAN (mod 10, norme GS1), contrôle de cohérence entre prix affiché et prix
  recalculé, logique des paliers dégressifs. C'est le fichier par lequel entrer.
- `extracteurs/generique.py` : extracteur de repli, qui détecte les colonnes par
  leur intitulé plutôt que par leur position.
- `extracteurs/pdf_base.py`, `extracteurs/paliers_lignes.py` : socles spécialisés.
- `extracteurs/fournisseur_a.py`, `fournisseur_b_conditions.py`, `fournisseur_c.py` :
  trois exemples conservés sur la soixantaine existante. Le deuxième est le plus
  instructif : il lit des remises composées, où c'est le format de la cellule et non
  sa valeur qui distingue un prix d'un taux.

**2. Périmètre.** Le champ de travail est figé une fois, puis lu, jamais recalculé.

- `perimetre_config.py` : les règles d'entrée, écrites comme des décisions métier
  assumées et non comme des paramètres à optimiser.
- `figer_perimetre.py` : produit la liste figée et son manifeste (empreintes des
  sources).
- `perimetre_liste.py` : la lit, et refuse de continuer si le compte ne tombe pas.
- `perimetre.py`, `normalisation.py` : annotation et normalisation des références.

**3. Rapprochement en cascade.** Plusieurs clés testées par fiabilité décroissante,
la clé retenue étant toujours tracée en colonne.

- `pa_completer.py` : le cœur. Quatre clés automatiques, garde-fou anti-aberration.
- `pa_identifiants.py` : cinquième passe, sur les identifiants abîmés.
- `pa_correspondances.py` : sixième et dernière, une table de décisions humaines qui
  ne comble que les cases vides.
- `pa_libelle.py` et `par_libelle.py` : rapprochement approximatif par libellé, à
  seuil calibré, en lecture seule.
- `pa_par_modele.py` : rapprochement par nom de modèle, quand ni la référence ni le
  libellé ne suffisent.
- `pa_conditionnement.py` : ramène un prix au lot vers un prix unitaire, jamais sans
  corroboration.

**4. Contrôle et restitution.**

- `qualite.py` : construction et hiérarchisation des anomalies.
- `pa_reconcilier_responsable.py` : confronte notre lecture à celle d'un second
  pipeline, indépendant, qui lit les mêmes tarifs.
- `ean_controle.py` : quatre niveaux de contrôle sur les codes retenus, en lecture
  seule.
- `controle_calculs.py` : déroule la chaîne de calcul ligne à ligne, sans rien écrire.
- `pa_fusionner.py`, `pa_etat_par_article.py`, `pa_ecartes.py`, `pa_a_arbitrer.py` :
  fusion non destructive et vues de relecture.
- `mise_en_forme.py` : mise en forme des classeurs, et écriture qui ne casse pas un
  traitement long quand le fichier cible est ouvert ailleurs.

**5. Chaîne EAN et sources externes.**

- `decoder_gs1.py` : décodage GS1-128 et DataMatrix, avec distinction entre unité de
  vente et carton.
- `ean_global.py` : consolidation multi-sources, avec ordre de priorité explicite.
- `ean_version.py` : gel d'une version numérotée, arbitrage des codes ambigus.
- `eudamed.py`, `eudamed_par_reference.py`, `eudamed_fiabilite.py`,
  `ean_verifier_libelle.py` : interrogation du registre européen public des
  dispositifs médicaux, et mesure de la confiance à accorder à ses réponses.
- `ar/base.py`, `ar/fournisseur_d.py`, `extraire_ar.ps1` : lecture des accusés de
  réception fournisseurs, qui donnent des prix réellement facturés, à une date.

**Orchestration.** `run_pa.py` enchaîne les étapes, s'arrête au premier échec
bloquant, et refuse de démarrer si les données de référence sont incohérentes.

## Points d'ingénierie notables

- **Garde-fous anti-aberration.** Au-delà d'un certain rapport avec le prix de
  référence, une valeur n'est pas injectée mais déviée vers une colonne de
  vérification. Un écart d'un ordre de grandeur n'est pas une négociation
  commerciale, c'est une erreur de lecture, et une erreur de lecture ne se signale
  pas d'elle-même.

- **Aucune suppression silencieuse.** La population complète entre dans le pipeline,
  et tout filtre devient une colonne. Une ligne écartée porte son motif et reste
  consultable. Un filtre appliqué en amont ne se voit que dans le total : on ne sait
  plus ce qui a été retiré, ni pourquoi, ni combien.

- **Calibration mesurée, pas au jugé.** Le seuil du rapprochement par libellé a été
  réglé sur un jeu d'épreuve construit à partir de correspondances déjà connues, en
  masquant la référence pour ne garder que le libellé. La mesure a montré que le
  seuil envisagé au départ était inatteignable sur les cas qui comptaient.

- **Détection d'un bug silencieux de conversion d'identifiants.** Une bibliothèque de
  traitement de données convertit une colonne d'identifiants en nombres à virgule dès
  qu'une valeur manque. La jointure tombe alors à zéro sans lever la moindre erreur.
  La normalisation correspondante est appliquée de façon systématique.

- **Réconciliation contre un pipeline tiers.** Un second pipeline, écrit par
  quelqu'un d'autre, lit les mêmes tarifs. Les deux lectures sont confrontées
  automatiquement. Une première version de ce contrôle comparait, à cause d'une clé
  de jointure trop stricte, une fraction infime des lignes tout en concluant à
  l'absence de divergence : un feu vert non mérité, qui est le pire résultat possible.

## Ce qui n'est pas ici, et pourquoi

- **Les données.** Tarifs fournisseurs, exports du système d'entrepôt, classeurs de
  travail, accusés de réception : rien de tout cela n'est publié. Le `.gitignore`
  exclut les formats correspondants.

- **Le registre des conditions commerciales.** Le module central qui associe chaque
  fournisseur à son dossier de tarifs, à ses remises négociées et à ses
  particularités de lecture n'est pas inclus : c'est de l'information commerciale.
  Il est importé par une bonne partie des fichiers présents, ce qui explique une
  partie des imports non résolus.

- **L'application de relevé terrain.** Le serveur et l'interface de scan de
  codes-barres en entrepôt (lecture par caméra ou par douchette USB) ont été sortis
  dans un dépôt séparé, leur cycle de vie n'étant pas celui du pipeline.

- **La chaîne OCR.** Certains tarifs n'existent qu'en PDF image et demandent un OCR.
  Le module qui s'en charge ne fait pas partie de cet extrait ; les fichiers présents
  ici traitent des PDF à couche texte. Les limites connues de cette approche sont
  documentées dans le code : un caractère mal reconnu ne produit pas une erreur mais
  un prix faux et silencieux, ce qui impose de croiser deux moteurs et de recouper
  par un contrôle de vraisemblance.

## Stack technique

Python 3, `pandas` et `numpy` pour le traitement tabulaire, `openpyxl` pour la
lecture et l'écriture de classeurs Excel, `pdfplumber` pour les PDF à couche texte,
`rapidfuzz` pour le rapprochement approximatif de libellés, `requests` pour
l'interrogation d'API publiques. Un script PowerShell pour l'extraction de pièces
jointes depuis une messagerie locale.

## Contexte

Projet conçu et réalisé en autonomie dans le cadre d'une mission achats et supply
chain, sur plusieurs milliers de références et près de 90 fournisseurs. La couverture
en codes EAN exploitables a été portée à plus de 70 % du périmètre suivi, et la
couverture en prix d'achat à jour à plus de 75 %.

## Licence

Code fourni à titre de démonstration. Tous droits réservés.
