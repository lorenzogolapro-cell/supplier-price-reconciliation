# Extraction of the supplier order acknowledgements from Outlook.
#
#   powershell -ExecutionPolicy Bypass -File extraire_ar.ps1
#   powershell -ExecutionPolicy Bypass -File extraire_ar.ps1 -MoisMax 6 -ParFournisseur 200
#
# Read only: we read the messages and save a copy of the attachments. No
# message is modified, moved or marked as read.
#
# The time window matters: an acknowledgement from 2023 says nothing
# about today's price. -MoisMax bounds the extraction to the useful
# history.
#
# Two traps handled here:
#   - at supplier B every attachment carries the same name. Without
#     renaming by acknowledgement number, each file overwrites the
#     previous one.
#   - PowerShell 5.1 reads a .ps1 without a BOM as ANSI: no accented
#     literal in the folder navigation, or nothing is ever found. Hence
#     the pattern comparisons ('te de r').

param(
    [int]$MoisMax = 18,
    [int]$ParFournisseur = 500,
    [string]$Destination = "$PSScriptRoot\ar_pdf",
    # Limits the pass to a single supplier (pattern on the folder name).
    # Without it, adding one rule means re-extracting the seven others.
    [string]$Fournisseur = ''
)

# One pattern per supplier. Subject and attachment serve to rule out what
# is not an acknowledgement: shipping notices, delivery notes, mail
# threads. The Expediteur field is the main safeguard: most of these
# folders ALSO hold our own purchase orders, sent by our internal
# Exchange (/O=EXCHANGELABS/...). Those carry OUR prices taken from the
# WMS - reading them would teach us nothing. Only mail coming from the
# supplier carries the price it commits to invoicing.
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
    # Very clean acknowledgement, but with no reference code: the
    # matching will have to go through the label and the size.
    @{ Dossier = 'FournisseurG'
       Expediteur = 'fournisseur-cb\.example|fournisseur-g'
       Sujet   = '.'
       PJ      = '^AR_CMD_.*\.pdf$'
       NumAR   = 'Commande\s+(\d+)'
       NumCde  = $null }
    # Carries the GROSS price, the contractual discount and the NET price
    # - all three, which no price list gives.
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
    # Supplier A, added on 21/09. The Outlook folder name carries an
    # accent: hence MotifDossier, because an accented literal cannot be
    # compared in a .ps1 without a BOM.
    #
    # The sorting is clear in this 690-message folder:
    #   supplier AR  -> supplier service address,
    #                   subject "1CDE...", attachment "1CDE....pdf"
    #   our orders   -> internal Exchange, subject "Commande CMD-EXEMPLE-1"
    # The attachment pattern is therefore enough to rule out our own
    # purchase orders and the mail threads (which carry nothing but
    # signature images).
    @{ Dossier = 'FournisseurA'
       MotifDossier = '^FournisseurA'
       Expediteur = 'fournisseur-a\.example'
       Sujet   = '1CDE\d+'
       PJ      = '^1CDE\d+\.pdf$'
       NumAR   = '(1CDE\d+)'
       NumCde  = $null }
    # Supplier D, added on 21/09.
    #
    # The folder mixes TWO automatic mails from the same domain:
    #   commandes@   -> the acknowledgement, with prices  <- what we want
    #   expeditions@ -> the shipping notice, with none
    # Hence a filter on the full address and not on the domain.
    #
    # TRAP: the attachment is ALWAYS called
    # "Order_Acknowledgement.pdf". Without renaming by acknowledgement
    # number, each file overwrites the previous one - exactly the
    # supplier B case documented above. The NumAR taken from the subject
    # is therefore mandatory.
    #
    # The accents of the subject are replaced by dots: an accented
    # literal in a .ps1 without a BOM matches nothing (see the header).
    @{ Dossier = 'FournisseurD'
       MotifDossier = '^FournisseurD'
       Expediteur = 'commandes@fournisseur-d\.example'
       Sujet   = 'Accus.* de r.ception de commande num.ro'
       PJ      = '^Order_Acknowledgement\.pdf$'
       NumAR   = 'num.ro\s+(\d+)'
       NumCde  = $null }
    # Supplier E, added on 21/09.
    #
    # commandes@fournisseur-e.example ALSO forwards our own purchase
    # orders ("TR: commande CMD-EXEMPLE-2"): the attachment is what
    # tells them apart, the acknowledgement's one ending with
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
    Write-Output "limited to supplier: $Fournisseur ($($Regles.Count) rule(s))"
}

$Limite = (Get-Date).AddMonths(-$MoisMax)
Write-Output "acknowledgements received since $($Limite.ToString('dd/MM/yyyy'))"

try {
    $ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
} catch {
    $ol = New-Object -ComObject Outlook.Application
}
$ns = $ol.GetNamespace('MAPI')

$store = $null
foreach ($s in $ns.Folders) { if ($s.Name -like 'Achats*') { $store = $s; break } }
if (-not $store) { throw "Shared mailbox 'Achats' not found" }

$inbox = $null
foreach ($f in $store.Folders) { if ($f.Name -match 'te de r') { $inbox = $f; break } }
if (-not $inbox) { throw "Inbox not found" }

$racine = $null
foreach ($f in $inbox.Folders) { if ($f.Name -eq 'FOURNISSEURS') { $racine = $f; break } }
if (-not $racine) { throw "FOURNISSEURS folder not found" }

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
    if (-not $cible) { Write-Output "$nom : folder not found"; continue }

    $items = $cible.Items
    $items.Sort('[ReceivedTime]', $true)

    $pris = 0
    for ($i = 1; $i -le $items.Count -and $pris -lt $ParFournisseur; $i++) {
        $m = $items.Item($i)
        if ($m.Class -ne 43) { continue }
        # Sorted by descending date: past the limit, everything else is
        # older still.
        if ($m.ReceivedTime -lt $Limite) { break }
        if ($m.Subject -notmatch $regle.Sujet) { continue }
        # Rules out our own purchase orders, sent by our own Exchange
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
    Write-Output "$nom : $pris acknowledgements extracted"
}

# The index is APPENDED to, it is never overwritten.
#
# With -Fournisseur, the pass only sees one folder: writing the index as
# it stands would erase the rows of the other suppliers extracted before.
# So we re-read what exists, drop the rows of the supplier(s) we have
# just redone (they are replaced, not duplicated), and merge.
$csv = Join-Path $Destination 'index_ar.csv'
$refaits = @($Regles | ForEach-Object { $_.Dossier })
$garde = @()
if (Test-Path $csv) {
    $ancien = @(Import-Csv -Path $csv)
    $garde = @($ancien | Where-Object { $refaits -notcontains $_.Fournisseur })
    Write-Output ""
    Write-Output "existing index: $($ancien.Count) rows, of which $($garde.Count) kept"
}
$tout = @($garde) + @($index)
$tout | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
Write-Output ""
Write-Output "$($index.Count) attachments extracted -> $Destination"
Write-Output "index: $($tout.Count) rows in total"
