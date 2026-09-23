# Extraction des accuses de reception fournisseurs depuis Outlook.
#
#   powershell -ExecutionPolicy Bypass -File extraire_ar.ps1
#   powershell -ExecutionPolicy Bypass -File extraire_ar.ps1 -MoisMax 6 -ParFournisseur 200
#
# Lecture seule : on lit les messages et on enregistre une copie des
# pieces jointes. Aucun message n'est modifie, deplace ni marque lu.
#
# La fenetre de temps compte : un AR de 2023 ne dit rien du prix
# d'aujourd'hui. -MoisMax borne l'extraction a l'historique utile.
#
# Deux pieges traites ici :
#   - chez le fournisseur B toutes les PJ portent le meme nom. Sans
#     renommage par numero d'AR, chaque fichier ecrase le precedent.
#   - PowerShell 5.1 lit un .ps1 sans BOM en ANSI : aucun litteral
#     accentue dans la navigation de dossiers, sous peine de ne rien
#     trouver. D'ou les comparaisons par motif ('te de r').

param(
    [int]$MoisMax = 18,
    [int]$ParFournisseur = 500,
    [string]$Destination = "$PSScriptRoot\ar_pdf",
    # Limite la passe a un seul fournisseur (motif sur le nom de dossier).
    # Sans cela, ajouter une regle oblige a re-extraire les sept autres.
    [string]$Fournisseur = ''
)

# Un motif par fournisseur. Sujet et PJ servent a ecarter ce qui n'est
# pas un AR : avis d'expedition, bons de livraison, echanges de mails.
# Le champ Expediteur est le garde-fou principal : la plupart de ces
# dossiers contiennent AUSSI nos propres bons de commande, envoyes par
# notre Exchange interne (/O=EXCHANGELABS/...). Ceux-la portent NOS prix
# tires du WMS - les lire n'apprendrait rien. Seul le courrier venu du
# fournisseur porte le prix qu'il s'engage a facturer.
$Regles = @(
    @{ Dossier = 'FournisseurB'
       Expediteur = 'fournisseur-b\.example'
       Sujet   = 'Confirmation commande Fournisseur B'
       PJ      = '\.pdf$'
       NumAR   = 'Fournisseur B\s+(\d+)'
       NumCde  = $null }
    @{ Dossier = 'FournisseurH'
       Expediteur = 'fournisseur-h\.example'
       Sujet   = 'Accus'
       PJ      = '^AR_H_.*\.pdf$'
       NumAR   = 'Commande\s+(A\d+)'
       NumCde  = 'Votre r.f\.\s*(\d+)' }
    @{ Dossier = 'FournisseurL'
       Expediteur = 'fournisseur-l\.example|fournisseur-l-sav'
       Sujet   = 'Votre Commande FOURNISSEUR L'
       PJ      = '^FOURNISSEUR_L_Commande_.*\.pdf$'
       NumAR   = 'n.(\d+)'
       NumCde  = 'r.f.rence\s+(\d+)' }
    # Accuse tres propre, mais sans code reference : le rapprochement
    # devra passer par le libelle et la taille.
    @{ Dossier = 'FournisseurG'
       Expediteur = 'fournisseur-cb\.example|fournisseur-g'
       Sujet   = '.'
       PJ      = '^AR_CMD_.*\.pdf$'
       NumAR   = 'Commande\s+(\d+)'
       NumCde  = $null }
    # Porte le prix BRUT, la remise de convention et le prix NET -
    # les trois, ce qu'aucun tarif ne donne.
    @{ Dossier = 'FournisseurN'
       Expediteur = 'fournisseur-n\.example'
       Sujet   = '.'
       PJ      = '^\d+\.pdf$'
       NumAR   = $null
       NumCde  = 'commande\s+(\d{9})' }
    @{ Dossier = 'FournisseurM'
       Expediteur = 'fournisseur-m\.example'
       Sujet   = '.'
       PJ      = '^Commande_C\d+\.pdf$'
       NumAR   = 'Commande_(C\d+)'
       NumCde  = 'commande\s+(\d{9})' }
    @{ Dossier = 'FournisseurK'
       Expediteur = 'fournisseur-k\.example'
       Sujet   = 'Bon de commande'
       PJ      = '^Bon_Commande_.*\.pdf$'
       NumAR   = 'N°?\s*(C\d+)'
       NumCde  = $null }
    # Fournisseur A, ajoute le 21/09. Le dossier Outlook porte un accent
    # dans son nom : d'ou MotifDossier, parce qu'un litteral accentue ne
    # se compare pas dans un .ps1 sans BOM.
    #
    # Le tri est net dans ce dossier de 690 messages :
    #   AR du fournisseur -> adresse de service du fournisseur,
    #                        sujet "1CDE...", piece jointe "1CDE....pdf"
    #   nos commandes     -> Exchange interne, sujet "Commande CMD-EXEMPLE-1"
    # Le motif de PJ suffit donc a ecarter nos propres bons de commande et
    # les fils de discussion (qui ne portent que des images de signature).
    @{ Dossier = 'FournisseurA'
       MotifDossier = '^FournisseurA'
       Expediteur = 'fournisseur-a\.example'
       Sujet   = '1CDE\d+'
       PJ      = '^1CDE\d+\.pdf$'
       NumAR   = '(1CDE\d+)'
       NumCde  = $null }
    # Fournisseur D, ajoute le 21/09.
    #
    # Le dossier melange DEUX courriers automatiques du meme domaine :
    #   commandes@   -> l'AR, avec les prix   <- ce qu'on veut
    #   expeditions@ -> l'avis d'expedition, sans prix
    # D'ou un filtre sur l'adresse complete et pas sur le domaine.
    #
    # PIEGE : la piece jointe s'appelle TOUJOURS
    # "Order_Acknowledgement.pdf". Sans renommage par numero d'AR,
    # chaque fichier ecrase le precedent - exactement le cas du
    # fournisseur B documente plus haut. Le NumAR tire du sujet est donc
    # obligatoire.
    #
    # Les accents du sujet sont remplaces par des points : un litteral
    # accentue dans un .ps1 sans BOM ne matche rien (voir l'en-tete).
    @{ Dossier = 'FournisseurD'
       MotifDossier = '^FournisseurD'
       Expediteur = 'commandes@fournisseur-d\.example'
       Sujet   = 'Accus.* de r.ception de commande num.ro'
       PJ      = '^Order_Acknowledgement\.pdf$'
       NumAR   = 'num.ro\s+(\d+)'
       NumCde  = $null }
    # Fournisseur E, ajoute le 21/09.
    #
    # commandes@fournisseur-e.example envoie AUSSI nos propres bons de
    # commande en transfert ("TR: commande CMD-EXEMPLE-2") : c'est la
    # piece jointe qui departage, celle de l'AR se terminant par
    # "Fournisseur E.pdf".
    @{ Dossier = 'FournisseurE'
       MotifDossier = '^FournisseurE'
       Expediteur = 'fournisseur-e\.example'
       Sujet   = 'Commande\s+\d+\s+Fournisseur E'
       PJ      = 'Fournisseur E\.pdf$'
       NumAR   = 'Commande\s+(\d+)\s+Fournisseur E'
       NumCde  = $null }
)

if ($Fournisseur) {
    $Regles = @($Regles | Where-Object { $_.Dossier -match $Fournisseur })
    Write-Output "limite au fournisseur : $Fournisseur ($($Regles.Count) regle(s))"
}

$Limite = (Get-Date).AddMonths(-$MoisMax)
Write-Output "AR recus depuis le $($Limite.ToString('dd/MM/yyyy'))"

try {
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch {
    $ol = New-Object -ComObject Outlook.Application
}
$ns = $ol.GetNamespace('MAPI')

$store = $null
foreach ($s in $ns.Folders) { if ($s.Name -like 'Achats*') { $store = $s; break } }
if (-not $store) { throw "Boite partagee 'Achats' introuvable" }

$inbox = $null
foreach ($f in $store.Folders) { if ($f.Name -match 'te de r') { $inbox = $f; break } }
if (-not $inbox) { throw "Boite de reception introuvable" }

$racine = $null
foreach ($f in $inbox.Folders) { if ($f.Name -eq 'FOURNISSEURS') { $racine = $f; break } }
if (-not $racine) { throw "Dossier FOURNISSEURS introuvable" }

$index = @()

foreach ($regle in $Regles) {
    $nom = $regle.Dossier
    $dossier = Join-Path $Destination $nom
    New-Item -ItemType Directory -Force -Path $dossier | Out-Null

    $cible = $null
    $motif = $regle.MotifDossier
    foreach ($f in $racine.Folders) {
        if ($motif) { if ($f.Name -match $motif) { $cible = $f; break } }
        elseif ($f.Name -eq $nom) { $cible = $f; break }
    }
    if (-not $cible) { Write-Output "$nom : dossier absent"; continue }

    $items = $cible.Items
    $items.Sort('[ReceivedTime]', $true)

    $pris = 0
    for ($i = 1; $i -le $items.Count -and $pris -lt $ParFournisseur; $i++) {
        $m = $items.Item($i)
        if ($m.Class -ne 43) { continue }
        # Trie par date decroissante : passe la limite, tout le reste
        # est plus ancien encore.
        if ($m.ReceivedTime -lt $Limite) { break }
        if ($m.Subject -notmatch $regle.Sujet) { continue }
        # Ecarte nos propres bons de commande, envoyes par notre Exchange
        if ($regle.Expediteur) {
            $de = "$($m.SenderEmailAddress) $($m.SenderName)"
            if ($de -notmatch $regle.Expediteur) { continue }
        }

        $numAR = ''
        if ($regle.NumAR -and $m.Subject -match $regle.NumAR) { $numAR = $Matches[1] }
        $numCde = ''
        if ($regle.NumCde -and $m.Subject -match $regle.NumCde) { $numCde = $Matches[1] }

        $garde = $false
        foreach ($pj in $m.Attachments) {
            if ($pj.FileName -notmatch $regle.PJ) { continue }
            $ext = [IO.Path]::GetExtension($pj.FileName)
            $base = if ($numAR) { "$nom`_$numAR" }
                    else { "$nom`_$($m.ReceivedTime.ToString('yyyyMMdd_HHmmss'))" }
            $fichier = Join-Path $dossier "$base$ext"
            $n = 2
            while (Test-Path $fichier) { $fichier = Join-Path $dossier "$base`_$n$ext"; $n++ }
            $pj.SaveAsFile($fichier)
            $garde = $true

            $index += [pscustomobject]@{
                Fournisseur  = $nom
                Recu         = $m.ReceivedTime.ToString('yyyy-MM-dd HH:mm:ss')
                NumAR        = $numAR
                CommandeWms  = $numCde
                Expediteur   = $m.SenderEmailAddress
                Sujet        = $m.Subject
                Fichier      = Split-Path $fichier -Leaf
            }
        }
        if ($garde) { $pris++ }
    }
    Write-Output "$nom : $pris AR extraits"
}

# L'index se COMPLETE, il ne s'ecrase pas.
#
# Avec -Fournisseur, la passe ne voit qu'un seul dossier : ecrire l'index
# tel quel effacerait les lignes des autres fournisseurs extraits avant.
# On relit donc l'existant, on retire les lignes du ou des fournisseurs
# qu'on vient de refaire (elles sont remplacees, pas dupliquees), et on
# rassemble.
$csv = Join-Path $Destination 'index_ar.csv'
$refaits = @($Regles | ForEach-Object { $_.Dossier })
$garde = @()
if (Test-Path $csv) {
    $ancien = @(Import-Csv -Path $csv)
    $garde = @($ancien | Where-Object { $refaits -notcontains $_.Fournisseur })
    Write-Output ""
    Write-Output "index existant : $($ancien.Count) lignes, dont $($garde.Count) conservees"
}
$tout = @($garde) + @($index)
$tout | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
Write-Output ""
Write-Output "$($index.Count) pieces jointes extraites -> $Destination"
Write-Output "index : $($tout.Count) lignes au total"
