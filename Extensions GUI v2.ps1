#Requires -Version 5.1
<#
.SYNOPSIS
    Interface graphique de comparaison de répertoire téléphonique
.DESCRIPTION
    Compare les données AD avec un fichier Excel et affiche les différences
#>

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ============================================
# CONFIGURATION
# ============================================

$OUPath = "OU=Users,OU=DFL-MTL,OU=Divisions,DC=groupedeschenes,DC=loc"

$locationMapping = @{
    "Sherbrooke-J1H" = "Sherbrooke"
    "Sherbrooke-J1L" = "Sherbrooke"
    "Quebec-G1V" = "Québec"
    "Montreal-H1Y" = "Montréal"
    "Laval-H7L" = "Laval"
    "Drummondville-J2C" = "Drummondville"
    "Granby-J2G" = "Granby"
}

# Variables globales
$script:adData = $null
$script:fileData = $null
$script:adDataCache = $null
$script:adDataCacheTime = $null
$script:cacheValidityMinutes = 5
$script:lastComparison = $null

# Couleurs
$colorPrimary = [System.Drawing.Color]::FromArgb(0, 120, 215)
$colorSecondary = [System.Drawing.Color]::FromArgb(243, 242, 241)
$colorSuccess = [System.Drawing.Color]::FromArgb(16, 124, 16)
$colorWarning = [System.Drawing.Color]::FromArgb(255, 185, 0)
$colorDanger = [System.Drawing.Color]::FromArgb(232, 17, 35)

# ============================================
# FONCTIONS
# ============================================

function Normalize-PhoneExtension {
    param([string]$Extension)

    if ([string]::IsNullOrWhiteSpace($Extension)) {
        return ""
    }

    # Normaliser: enlever espaces, tirets, parenthèses
    return $Extension -replace '[\s\-\(\)]', ''
}

function New-CustomDataGrid {
    param(
        [int]$X,
        [int]$Y,
        [int]$Width,
        [int]$Height,
        [string[]]$Columns,
        [System.Drawing.Color]$BackgroundColor = [System.Drawing.Color]::White
    )

    $dataGrid = New-Object System.Windows.Forms.DataGridView
    $dataGrid.Location = New-Object System.Drawing.Point($X, $Y)
    $dataGrid.Size = New-Object System.Drawing.Size($Width, $Height)
    $dataGrid.AllowUserToAddRows = $false
    $dataGrid.AllowUserToDeleteRows = $false
    $dataGrid.ReadOnly = $true
    $dataGrid.SelectionMode = 'FullRowSelect'
    $dataGrid.AutoSizeColumnsMode = 'Fill'
    $dataGrid.BackgroundColor = $BackgroundColor

    foreach ($colName in $Columns) {
        $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
        $col.HeaderText = $colName
        $col.Name = $colName
        [void]$dataGrid.Columns.Add($col)
    }

    return $dataGrid
}

function Get-AllPhoneDirectory {
    param(
        [string]$OUSearchBase,
        [hashtable]$LocationMap
    )
    
    try {
        $users = Get-ADUser -SearchBase $OUSearchBase -Filter {Enabled -eq $true} `
            -Properties GivenName, Surname, ipPhone, telephoneNumber, Department, 
                        EmailAddress, Title, OfficePhone, Mobile, Company, Office, 
                        City, l, PostalCode, StreetAddress, Manager, 
                        physicalDeliveryOfficeName |
            Where-Object { 
                $_.SamAccountName -notlike "*admin*" -and 
                $_.SamAccountName -notlike "*service*" -and
                $_.SamAccountName -notlike "*svc*" -and
                $_.ObjectClass -eq "user"
            }
        
        $allUsers = @()
        
        foreach ($user in $users) {
            $extension = if ($user.ipPhone) { $user.ipPhone } else { $user.telephoneNumber }
            $address = $user.StreetAddress
            
            # FILTRE: Exclure si PAS d'extension ET PAS d'adresse
            if ([string]::IsNullOrWhiteSpace($extension) -and [string]::IsNullOrWhiteSpace($address)) {
                continue
            }
            
            $city = $user.City
            if (-not $city) { $city = $user.l }
            if (-not $city) { $city = $user.Office }
            
            $postalCode = $user.PostalCode
            $postalPrefix = if ($postalCode -and $postalCode.Length -ge 3) {
                $postalCode.Substring(0,3).ToUpper()
            } else {
                ""
            }
            
            $branch = "Non specifie"
            if ($city -and $postalPrefix) {
                $locationKey = "$city-$postalPrefix"
                if ($LocationMap.ContainsKey($locationKey)) {
                    $branch = $LocationMap[$locationKey]
                } else {
                    $branch = $city
                }
            } elseif ($city) {
                $branch = $city
            }
            
            $managerName = ""
            if ($user.Manager) {
                try {
                    $manager = Get-ADUser -Identity $user.Manager -Properties DisplayName -ErrorAction SilentlyContinue
                    if ($manager) {
                        $managerName = $manager.DisplayName
                    }
                } catch {}
            }
            
            $allUsers += [PSCustomObject]@{
                Succursale = $branch
                Nom = if ($user.Surname) { $user.Surname } else { "" }
                Prenom = if ($user.GivenName) { $user.GivenName } else { "" }
                Adresse = if ($address) { $address } else { "" }
                Ville = if ($city) { $city } else { "" }
                CodePostal = if ($postalCode) { $postalCode } else { "" }
                Extension = if ($extension) { $extension } else { "" }
                Email = if ($user.EmailAddress) { $user.EmailAddress } else { "" }
                SamAccountName = $user.SamAccountName
            }
        }
        
        return $allUsers | Sort-Object Nom, Prenom
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur AD: $($_.Exception.Message)", "Erreur")
        return $null
    }
}

function Load-ExcelFile {
    param([string]$FilePath)
    
    try {
        Get-Process excel -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 1
        
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($FilePath)
        
        $worksheet = $null
        foreach ($sheet in $workbook.Worksheets) {
            if ($sheet.Name -like "*epertoire*" -or $sheet.Name -eq "Sheet1") {
                $worksheet = $sheet
                break
            }
        }
        if (-not $worksheet) {
            $worksheet = $workbook.Worksheets.Item(1)
        }
        
        $lastRow = $worksheet.UsedRange.Rows.Count
        $lastCol = $worksheet.UsedRange.Columns.Count
        
        # Trouver les colonnes
        $colMap = @{}
        for ($col = 1; $col -le $lastCol; $col++) {
            $header = $worksheet.Cells.Item(1, $col).Text
            $colMap[$header] = $col
        }
        
        $fileUsers = @()
        for ($i = 2; $i -le $lastRow; $i++) {
            $sam = $worksheet.Cells.Item($i, $colMap["SamAccountName"]).Text
            if ($sam -and $sam -ne "") {
                $fileUsers += [PSCustomObject]@{
                    Succursale = $worksheet.Cells.Item($i, $colMap["Succursale"]).Text
                    Nom = $worksheet.Cells.Item($i, $colMap["Nom"]).Text
                    Prenom = $worksheet.Cells.Item($i, $colMap["Prenom"]).Text
                    Adresse = $worksheet.Cells.Item($i, $colMap["Adresse"]).Text
                    Ville = $worksheet.Cells.Item($i, $colMap["Ville"]).Text
                    CodePostal = $worksheet.Cells.Item($i, $colMap["Code Postal"]).Text
                    Extension = $worksheet.Cells.Item($i, $colMap["Extension"]).Text
                    Email = $worksheet.Cells.Item($i, $colMap["Email"]).Text
                    SamAccountName = $sam.Trim()
                }
            }
        }
        
        $workbook.Close($false)
        $excel.Quit()
        
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        return $fileUsers | Sort-Object Nom, Prenom
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("Erreur lecture Excel: $($_.Exception.Message)", "Erreur")
        return $null
    }
}

function Export-ComparisonResults {
    param(
        $Nouveaux,
        $Partis,
        $Modifications,
        [string]$OutputPath
    )

    try {
        # Créer un objet pour l'export
        $exportData = @()

        # Ajouter les nouveaux
        foreach ($user in $Nouveaux) {
            $exportData += [PSCustomObject]@{
                Type = "NOUVEAU"
                Succursale = $user.Succursale
                Nom = $user.Nom
                Prenom = $user.Prenom
                Extension = $user.Extension
                Adresse = $user.Adresse
                Ville = $user.Ville
                CodePostal = $user.CodePostal
                Email = $user.Email
                SamAccountName = $user.SamAccountName
                Changements = ""
            }
        }

        # Ajouter les partis
        foreach ($user in $Partis) {
            $exportData += [PSCustomObject]@{
                Type = "PARTI"
                Succursale = $user.Succursale
                Nom = $user.Nom
                Prenom = $user.Prenom
                Extension = $user.Extension
                Adresse = $user.Adresse
                Ville = $user.Ville
                CodePostal = $user.CodePostal
                Email = $user.Email
                SamAccountName = $user.SamAccountName
                Changements = ""
            }
        }

        # Ajouter les modifications
        foreach ($modif in $Modifications) {
            $exportData += [PSCustomObject]@{
                Type = "MODIFICATION"
                Succursale = $modif.NouvelleSuccursale
                Nom = $modif.Nom
                Prenom = $modif.Prenom
                Extension = $modif.NouvelleExtension
                Adresse = $modif.NouvelleAdresse
                Ville = $modif.NouvelleVille
                CodePostal = ""
                Email = $modif.NouvelEmail
                SamAccountName = $modif.SamAccountName
                Changements = $modif.Changements
            }
        }

        # Exporter en CSV avec UTF8 BOM pour Excel
        $exportData | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8

        [System.Windows.Forms.MessageBox]::Show(
            "Export réussi: $OutputPath`n`nTotal: $($exportData.Count) entrées",
            "Export terminé",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Information
        )
        return $true
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show(
            "Erreur lors de l'export: $($_.Exception.Message)",
            "Erreur",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Error
        )
        return $false
    }
}

function Compare-Data {
    param(
        $ADData,
        $FileData
    )

    if (-not $ADData -or -not $FileData) {
        return $null
    }

    # Créer des hashtables pour accès rapide (insensible à la casse)
    $adHash = @{}
    foreach ($user in $ADData) {
        $adHash[$user.SamAccountName.ToLower()] = $user
    }

    $fileHash = @{}
    foreach ($user in $FileData) {
        $fileHash[$user.SamAccountName.ToLower()] = $user
    }

    $nouveaux = @()
    $partis = @()
    $modifications = @()

    # Trouver les nouveaux (dans AD mais pas dans fichier)
    foreach ($adUser in $ADData) {
        $samLower = $adUser.SamAccountName.ToLower()
        if (-not $fileHash.ContainsKey($samLower)) {
            $nouveaux += $adUser
        }
    }

    # Trouver les partis (dans fichier mais pas dans AD)
    foreach ($fileUser in $FileData) {
        $samLower = $fileUser.SamAccountName.ToLower()
        if (-not $adHash.ContainsKey($samLower)) {
            $partis += $fileUser
        }
    }

    # Trouver les modifications (présent dans les deux mais avec des différences)
    foreach ($adUser in $ADData) {
        $samLower = $adUser.SamAccountName.ToLower()
        if ($fileHash.ContainsKey($samLower)) {
            $fileUser = $fileHash[$samLower]
            $changes = @()

            # Comparer Extension (normalisée)
            $adExtNorm = Normalize-PhoneExtension $adUser.Extension
            $fileExtNorm = Normalize-PhoneExtension $fileUser.Extension
            if ($adExtNorm -ne $fileExtNorm) {
                $changes += "Extension: '$($fileUser.Extension)' → '$($adUser.Extension)'"
            }

            # Comparer Adresse
            if ($adUser.Adresse -ne $fileUser.Adresse) {
                $changes += "Adresse: '$($fileUser.Adresse)' → '$($adUser.Adresse)'"
            }

            # Comparer Ville
            if ($adUser.Ville -ne $fileUser.Ville) {
                $changes += "Ville: '$($fileUser.Ville)' → '$($adUser.Ville)'"
            }

            # Comparer Succursale
            if ($adUser.Succursale -ne $fileUser.Succursale) {
                $changes += "Succursale: '$($fileUser.Succursale)' → '$($adUser.Succursale)'"
            }

            # Comparer Code Postal
            if ($adUser.CodePostal -ne $fileUser.CodePostal) {
                $changes += "Code Postal: '$($fileUser.CodePostal)' → '$($adUser.CodePostal)'"
            }

            # Comparer Email
            if ($adUser.Email -ne $fileUser.Email) {
                $changes += "Email: '$($fileUser.Email)' → '$($adUser.Email)'"
            }

            # Si des changements existent, ajouter à la liste
            if ($changes.Count -gt 0) {
                $modifications += [PSCustomObject]@{
                    SamAccountName = $adUser.SamAccountName
                    Nom = $adUser.Nom
                    Prenom = $adUser.Prenom
                    Changements = $changes -join " | "
                    AncienneExtension = $fileUser.Extension
                    NouvelleExtension = $adUser.Extension
                    AncienneAdresse = $fileUser.Adresse
                    NouvelleAdresse = $adUser.Adresse
                    AncienneVille = $fileUser.Ville
                    NouvelleVille = $adUser.Ville
                    AncienneSuccursale = $fileUser.Succursale
                    NouvelleSuccursale = $adUser.Succursale
                    AncienEmail = $fileUser.Email
                    NouvelEmail = $adUser.Email
                }
            }
        }
    }

    return @{
        Nouveaux = $nouveaux
        Partis = $partis
        Modifications = $modifications
    }
}

# ============================================
# INTERFACE GRAPHIQUE
# ============================================

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Comparateur de Repertoire Telephonique'
$form.Size = New-Object System.Drawing.Size(1400, 850)
$form.StartPosition = 'CenterScreen'
$form.BackColor = $colorSecondary
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

# Barre de progression globale
$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-Object System.Drawing.Point(10, 805)
$progressBar.Size = New-Object System.Drawing.Size(1360, 25)
$progressBar.Style = 'Continuous'
$progressBar.Visible = $false
$form.Controls.Add($progressBar)

# Titre
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Location = New-Object System.Drawing.Point(20, 10)
$lblTitle.Size = New-Object System.Drawing.Size(1350, 35)
$lblTitle.Text = 'COMPARATEUR DE REPERTOIRE TELEPHONIQUE'
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $colorPrimary
$lblTitle.TextAlign = 'MiddleCenter'
$form.Controls.Add($lblTitle)

# ===== PANEL GAUCHE - DONNÉES AD =====
$panelAD = New-Object System.Windows.Forms.Panel
$panelAD.Location = New-Object System.Drawing.Point(10, 60)
$panelAD.Size = New-Object System.Drawing.Size(670, 450)
$panelAD.BackColor = [System.Drawing.Color]::White
$panelAD.BorderStyle = 'FixedSingle'

$lblADTitle = New-Object System.Windows.Forms.Label
$lblADTitle.Location = New-Object System.Drawing.Point(10, 10)
$lblADTitle.Size = New-Object System.Drawing.Size(650, 30)
$lblADTitle.Text = 'DONNEES ACTIVE DIRECTORY'
$lblADTitle.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$lblADTitle.ForeColor = $colorPrimary
$lblADTitle.TextAlign = 'MiddleCenter'

$btnLoadAD = New-Object System.Windows.Forms.Button
$btnLoadAD.Location = New-Object System.Drawing.Point(225, 50)
$btnLoadAD.Size = New-Object System.Drawing.Size(220, 40)
$btnLoadAD.Text = 'CHARGER DEPUIS AD'
$btnLoadAD.BackColor = $colorPrimary
$btnLoadAD.ForeColor = [System.Drawing.Color]::White
$btnLoadAD.FlatStyle = 'Flat'
$btnLoadAD.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$btnLoadAD.Add_Click({
    $btnLoadAD.Enabled = $false
    $btnLoadAD.Text = "Chargement en cours..."
    $progressBar.Visible = $true
    $progressBar.Value = 0

    # Vérifier si le cache est valide
    $useCache = $false
    if ($script:adDataCache -and $script:adDataCacheTime) {
        $cacheAge = (Get-Date) - $script:adDataCacheTime
        if ($cacheAge.TotalMinutes -lt $script:cacheValidityMinutes) {
            $useCache = $true
            $script:adData = $script:adDataCache
            $progressBar.Value = 100
        }
    }

    # Charger depuis AD si pas de cache valide
    if (-not $useCache) {
        $progressBar.Value = 10
        $script:adData = Get-AllPhoneDirectory -OUSearchBase $OUPath -LocationMap $locationMapping
        $progressBar.Value = 80

        # Mettre à jour le cache
        if ($script:adData) {
            $script:adDataCache = $script:adData
            $script:adDataCacheTime = Get-Date
        }
        $progressBar.Value = 100
    }

    if ($script:adData) {
        $dataGridAD.Rows.Clear()
        foreach ($user in $script:adData) {
            [void]$dataGridAD.Rows.Add(
                $user.Succursale, $user.Nom, $user.Prenom, $user.Extension,
                $user.Ville, $user.Email, $user.SamAccountName
            )
        }
        $cacheStatus = if ($useCache) { " (depuis cache)" } else { "" }
        $lblADCount.Text = "Total: $($script:adData.Count) utilisateurs$cacheStatus"
        $lblADCount.ForeColor = $colorSuccess
        $btnExportAD.Enabled = $true

        # Si les deux sont chargés, comparer automatiquement
        if ($script:fileData) {
            Compare-AndDisplay
        }
    }

    $progressBar.Visible = $false
    $btnLoadAD.Enabled = $true
    $btnLoadAD.Text = "CHARGER DEPUIS AD"
})

$lblADCount = New-Object System.Windows.Forms.Label
$lblADCount.Location = New-Object System.Drawing.Point(10, 100)
$lblADCount.Size = New-Object System.Drawing.Size(650, 20)
$lblADCount.Text = 'Aucune donnee chargee'
$lblADCount.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$lblADCount.TextAlign = 'MiddleCenter'

# Filtres de recherche AD
$lblFilterAD = New-Object System.Windows.Forms.Label
$lblFilterAD.Location = New-Object System.Drawing.Point(10, 125)
$lblFilterAD.Size = New-Object System.Drawing.Size(80, 20)
$lblFilterAD.Text = 'Filtrer:'
$lblFilterAD.Font = New-Object System.Drawing.Font('Segoe UI', 8)

$txtFilterAD = New-Object System.Windows.Forms.TextBox
$txtFilterAD.Location = New-Object System.Drawing.Point(70, 123)
$txtFilterAD.Size = New-Object System.Drawing.Size(590, 20)
$txtFilterAD.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$txtFilterAD.Add_TextChanged({
    if ($script:adData) {
        $filterText = $txtFilterAD.Text.ToLower()
        $dataGridAD.Rows.Clear()
        foreach ($user in $script:adData) {
            $matchName = $user.Nom.ToLower().Contains($filterText) -or $user.Prenom.ToLower().Contains($filterText)
            $matchSucc = $user.Succursale.ToLower().Contains($filterText)
            $matchExt = $user.Extension.Contains($filterText)
            if ($matchName -or $matchSucc -or $matchExt -or $filterText -eq "") {
                [void]$dataGridAD.Rows.Add(
                    $user.Succursale, $user.Nom, $user.Prenom, $user.Extension,
                    $user.Ville, $user.Email, $user.SamAccountName
                )
            }
        }
    }
})

$cols = @("Succursale", "Nom", "Prenom", "Extension", "Ville", "Email", "SamAccountName")
$dataGridAD = New-CustomDataGrid -X 10 -Y 150 -Width 650 -Height 290 -Columns $cols

$panelAD.Controls.Add($lblADTitle)
$panelAD.Controls.Add($btnLoadAD)
$panelAD.Controls.Add($lblADCount)
$panelAD.Controls.Add($lblFilterAD)
$panelAD.Controls.Add($txtFilterAD)
$panelAD.Controls.Add($dataGridAD)
$form.Controls.Add($panelAD)

# ===== PANEL DROIT - FICHIER EXCEL =====
$panelFile = New-Object System.Windows.Forms.Panel
$panelFile.Location = New-Object System.Drawing.Point(690, 60)
$panelFile.Size = New-Object System.Drawing.Size(670, 450)
$panelFile.BackColor = [System.Drawing.Color]::White
$panelFile.BorderStyle = 'FixedSingle'

$lblFileTitle = New-Object System.Windows.Forms.Label
$lblFileTitle.Location = New-Object System.Drawing.Point(10, 10)
$lblFileTitle.Size = New-Object System.Drawing.Size(650, 30)
$lblFileTitle.Text = 'DONNEES FICHIER EXCEL'
$lblFileTitle.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$lblFileTitle.ForeColor = $colorSuccess
$lblFileTitle.TextAlign = 'MiddleCenter'

$btnLoadFile = New-Object System.Windows.Forms.Button
$btnLoadFile.Location = New-Object System.Drawing.Point(225, 50)
$btnLoadFile.Size = New-Object System.Drawing.Size(220, 40)
$btnLoadFile.Text = 'CHARGER FICHIER EXCEL'
$btnLoadFile.BackColor = $colorSuccess
$btnLoadFile.ForeColor = [System.Drawing.Color]::White
$btnLoadFile.FlatStyle = 'Flat'
$btnLoadFile.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$btnLoadFile.Add_Click({
    $openDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openDialog.Filter = 'Fichiers Excel (*.xlsx;*.xls)|*.xlsx;*.xls'
    $openDialog.Title = "Selectionnez le fichier Excel"

    if ($openDialog.ShowDialog() -eq 'OK') {
        $btnLoadFile.Enabled = $false
        $btnLoadFile.Text = "Chargement en cours..."
        $progressBar.Visible = $true
        $progressBar.Value = 10

        $script:fileData = Load-ExcelFile -FilePath $openDialog.FileName
        $progressBar.Value = 80

        if ($script:fileData) {
            $dataGridFile.Rows.Clear()
            foreach ($user in $script:fileData) {
                [void]$dataGridFile.Rows.Add(
                    $user.Succursale, $user.Nom, $user.Prenom, $user.Extension,
                    $user.Ville, $user.Email, $user.SamAccountName
                )
            }
            $lblFileCount.Text = "Total: $($script:fileData.Count) utilisateurs"
            $lblFileCount.ForeColor = $colorSuccess
            $btnExportFile.Enabled = $true

            # Si les deux sont chargés, comparer automatiquement
            if ($script:adData) {
                Compare-AndDisplay
            }
        }

        $progressBar.Value = 100
        $progressBar.Visible = $false
        $btnLoadFile.Enabled = $true
        $btnLoadFile.Text = "CHARGER FICHIER EXCEL"
    }
})

$lblFileCount = New-Object System.Windows.Forms.Label
$lblFileCount.Location = New-Object System.Drawing.Point(10, 100)
$lblFileCount.Size = New-Object System.Drawing.Size(650, 20)
$lblFileCount.Text = 'Aucune donnee chargee'
$lblFileCount.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$lblFileCount.TextAlign = 'MiddleCenter'

# Filtres de recherche File
$lblFilterFile = New-Object System.Windows.Forms.Label
$lblFilterFile.Location = New-Object System.Drawing.Point(10, 125)
$lblFilterFile.Size = New-Object System.Drawing.Size(80, 20)
$lblFilterFile.Text = 'Filtrer:'
$lblFilterFile.Font = New-Object System.Drawing.Font('Segoe UI', 8)

$txtFilterFile = New-Object System.Windows.Forms.TextBox
$txtFilterFile.Location = New-Object System.Drawing.Point(70, 123)
$txtFilterFile.Size = New-Object System.Drawing.Size(590, 20)
$txtFilterFile.Font = New-Object System.Drawing.Font('Segoe UI', 8)
$txtFilterFile.Add_TextChanged({
    if ($script:fileData) {
        $filterText = $txtFilterFile.Text.ToLower()
        $dataGridFile.Rows.Clear()
        foreach ($user in $script:fileData) {
            $matchName = $user.Nom.ToLower().Contains($filterText) -or $user.Prenom.ToLower().Contains($filterText)
            $matchSucc = $user.Succursale.ToLower().Contains($filterText)
            $matchExt = $user.Extension.Contains($filterText)
            if ($matchName -or $matchSucc -or $matchExt -or $filterText -eq "") {
                [void]$dataGridFile.Rows.Add(
                    $user.Succursale, $user.Nom, $user.Prenom, $user.Extension,
                    $user.Ville, $user.Email, $user.SamAccountName
                )
            }
        }
    }
})

$dataGridFile = New-CustomDataGrid -X 10 -Y 150 -Width 650 -Height 290 -Columns $cols

$panelFile.Controls.Add($lblFileTitle)
$panelFile.Controls.Add($btnLoadFile)
$panelFile.Controls.Add($lblFileCount)
$panelFile.Controls.Add($lblFilterFile)
$panelFile.Controls.Add($txtFilterFile)
$panelFile.Controls.Add($dataGridFile)
$form.Controls.Add($panelFile)

# ===== PANEL BAS - DIFFERENCES =====
$panelDiff = New-Object System.Windows.Forms.Panel
$panelDiff.Location = New-Object System.Drawing.Point(10, 520)
$panelDiff.Size = New-Object System.Drawing.Size(1350, 230)
$panelDiff.BackColor = [System.Drawing.Color]::White
$panelDiff.BorderStyle = 'FixedSingle'

$lblDiffTitle = New-Object System.Windows.Forms.Label
$lblDiffTitle.Location = New-Object System.Drawing.Point(10, 10)
$lblDiffTitle.Size = New-Object System.Drawing.Size(1330, 30)
$lblDiffTitle.Text = 'DIFFERENCES'
$lblDiffTitle.Font = New-Object System.Drawing.Font('Segoe UI', 12, [System.Drawing.FontStyle]::Bold)
$lblDiffTitle.ForeColor = $colorWarning
$lblDiffTitle.TextAlign = 'MiddleCenter'

$lblDiffStats = New-Object System.Windows.Forms.Label
$lblDiffStats.Location = New-Object System.Drawing.Point(10, 50)
$lblDiffStats.Size = New-Object System.Drawing.Size(1330, 25)
$lblDiffStats.Text = 'Chargez les donnees pour voir les differences'
$lblDiffStats.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$lblDiffStats.TextAlign = 'MiddleCenter'

$tabControlDiff = New-Object System.Windows.Forms.TabControl
$tabControlDiff.Location = New-Object System.Drawing.Point(10, 85)
$tabControlDiff.Size = New-Object System.Drawing.Size(1330, 135)

# Onglet Nouveaux
$tabNouveaux = New-Object System.Windows.Forms.TabPage
$tabNouveaux.Text = 'NOUVEAUX EMPLOYES'
$tabNouveaux.BackColor = [System.Drawing.Color]::White

$dataGridNouveaux = New-CustomDataGrid -X 5 -Y 5 -Width 1315 -Height 100 -Columns $cols `
    -BackgroundColor ([System.Drawing.Color]::FromArgb(220, 255, 220))

$tabNouveaux.Controls.Add($dataGridNouveaux)
[void]$tabControlDiff.TabPages.Add($tabNouveaux)

# Onglet Partis
$tabPartis = New-Object System.Windows.Forms.TabPage
$tabPartis.Text = 'EMPLOYES PARTIS'
$tabPartis.BackColor = [System.Drawing.Color]::White

$dataGridPartis = New-CustomDataGrid -X 5 -Y 5 -Width 1315 -Height 100 -Columns $cols `
    -BackgroundColor ([System.Drawing.Color]::FromArgb(255, 220, 220))

$tabPartis.Controls.Add($dataGridPartis)
[void]$tabControlDiff.TabPages.Add($tabPartis)

# Onglet Modifications
$tabModifications = New-Object System.Windows.Forms.TabPage
$tabModifications.Text = 'MODIFICATIONS'
$tabModifications.BackColor = [System.Drawing.Color]::White

$colsModif = @("Nom", "Prenom", "SamAccountName", "Changements", "Ancien Ext", "Nouvel Ext", "Ancienne Ville", "Nouvelle Ville")
$dataGridModifications = New-CustomDataGrid -X 5 -Y 5 -Width 1315 -Height 100 -Columns $colsModif `
    -BackgroundColor ([System.Drawing.Color]::FromArgb(255, 248, 220))

$tabModifications.Controls.Add($dataGridModifications)
[void]$tabControlDiff.TabPages.Add($tabModifications)

$panelDiff.Controls.Add($lblDiffTitle)
$panelDiff.Controls.Add($lblDiffStats)
$panelDiff.Controls.Add($tabControlDiff)
$form.Controls.Add($panelDiff)

# ===== BOUTONS D'EXPORT (barre basse) =====

# Bouton Export Répertoire AD
$btnExportAD = New-Object System.Windows.Forms.Button
$btnExportAD.Location = New-Object System.Drawing.Point(10, 762)
$btnExportAD.Size = New-Object System.Drawing.Size(310, 35)
$btnExportAD.Text = 'EXPORTER REPERTOIRE AD (CSV)'
$btnExportAD.BackColor = $colorPrimary
$btnExportAD.ForeColor = [System.Drawing.Color]::White
$btnExportAD.FlatStyle = 'Flat'
$btnExportAD.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$btnExportAD.Enabled = $false
$btnExportAD.Add_Click({
    if ($script:adData) {
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = 'Fichiers CSV (*.csv)|*.csv'
        $saveDialog.Title = "Enregistrer le repertoire AD"
        $saveDialog.FileName = "Repertoire_AD_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($saveDialog.ShowDialog() -eq 'OK') {
            $script:adData | Select-Object Succursale, Nom, Prenom, Extension, Adresse, Ville, CodePostal, Email, SamAccountName |
                Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show(
                "Export réussi: $($saveDialog.FileName)`nTotal: $($script:adData.Count) utilisateurs",
                "Export AD terminé", 'OK', 'Information')
        }
    }
})
$form.Controls.Add($btnExportAD)

# Bouton Export Répertoire Fichier
$btnExportFile = New-Object System.Windows.Forms.Button
$btnExportFile.Location = New-Object System.Drawing.Point(330, 762)
$btnExportFile.Size = New-Object System.Drawing.Size(310, 35)
$btnExportFile.Text = 'EXPORTER REPERTOIRE FICHIER (CSV)'
$btnExportFile.BackColor = $colorSuccess
$btnExportFile.ForeColor = [System.Drawing.Color]::White
$btnExportFile.FlatStyle = 'Flat'
$btnExportFile.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$btnExportFile.Enabled = $false
$btnExportFile.Add_Click({
    if ($script:fileData) {
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = 'Fichiers CSV (*.csv)|*.csv'
        $saveDialog.Title = "Enregistrer le repertoire Fichier"
        $saveDialog.FileName = "Repertoire_Fichier_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($saveDialog.ShowDialog() -eq 'OK') {
            $script:fileData | Select-Object Succursale, Nom, Prenom, Extension, Adresse, Ville, CodePostal, Email, SamAccountName |
                Export-Csv -Path $saveDialog.FileName -NoTypeInformation -Encoding UTF8
            [System.Windows.Forms.MessageBox]::Show(
                "Export réussi: $($saveDialog.FileName)`nTotal: $($script:fileData.Count) utilisateurs",
                "Export Fichier terminé", 'OK', 'Information')
        }
    }
})
$form.Controls.Add($btnExportFile)

# Bouton Export Différences
$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Location = New-Object System.Drawing.Point(650, 762)
$btnExport.Size = New-Object System.Drawing.Size(310, 35)
$btnExport.Text = 'EXPORTER LES DIFFERENCES (CSV)'
$btnExport.BackColor = $colorWarning
$btnExport.ForeColor = [System.Drawing.Color]::White
$btnExport.FlatStyle = 'Flat'
$btnExport.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$btnExport.Enabled = $false
$btnExport.Add_Click({
    if ($script:lastComparison) {
        $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
        $saveDialog.Filter = 'Fichiers CSV (*.csv)|*.csv'
        $saveDialog.Title = "Enregistrer les resultats"
        $saveDialog.FileName = "Comparaison_Repertoire_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
        if ($saveDialog.ShowDialog() -eq 'OK') {
            Export-ComparisonResults `
                -Nouveaux $script:lastComparison.Nouveaux `
                -Partis $script:lastComparison.Partis `
                -Modifications $script:lastComparison.Modifications `
                -OutputPath $saveDialog.FileName
        }
    }
})
$form.Controls.Add($btnExport)

# Fonction de comparaison
function Compare-AndDisplay {
    $comparison = Compare-Data -ADData $script:adData -FileData $script:fileData

    if ($comparison) {
        # Nouveaux
        $dataGridNouveaux.Rows.Clear()
        foreach ($user in $comparison.Nouveaux) {
            [void]$dataGridNouveaux.Rows.Add(
                $user.Succursale, $user.Nom, $user.Prenom, $user.Extension,
                $user.Ville, $user.Email, $user.SamAccountName
            )
        }

        # Partis
        $dataGridPartis.Rows.Clear()
        foreach ($user in $comparison.Partis) {
            [void]$dataGridPartis.Rows.Add(
                $user.Succursale, $user.Nom, $user.Prenom, $user.Extension,
                $user.Ville, $user.Email, $user.SamAccountName
            )
        }

        # Modifications
        $dataGridModifications.Rows.Clear()
        foreach ($modif in $comparison.Modifications) {
            [void]$dataGridModifications.Rows.Add(
                $modif.Nom, $modif.Prenom, $modif.SamAccountName,
                $modif.Changements, $modif.AncienneExtension, $modif.NouvelleExtension,
                $modif.AncienneVille, $modif.NouvelleVille
            )
        }

        # Stats
        $nouveauxCount = $comparison.Nouveaux.Count
        $partisCount = $comparison.Partis.Count
        $modifCount = $comparison.Modifications.Count
        $lblDiffStats.Text = "NOUVEAUX: $nouveauxCount  |  PARTIS: $partisCount  |  MODIFICATIONS: $modifCount"

        if ($nouveauxCount -gt 0 -or $modifCount -gt 0) {
            $lblDiffStats.ForeColor = $colorWarning
        } elseif ($partisCount -gt 0) {
            $lblDiffStats.ForeColor = $colorDanger
        } else {
            $lblDiffStats.ForeColor = $colorPrimary
            $lblDiffStats.Text = "Aucune difference detectee"
        }

        # Activer le bouton export si des différences existent
        if ($nouveauxCount -gt 0 -or $partisCount -gt 0 -or $modifCount -gt 0) {
            $btnExport.Enabled = $true
        } else {
            $btnExport.Enabled = $false
        }

        # Stocker pour l'export
        $script:lastComparison = $comparison
    }
}

# Afficher le formulaire
[void]$form.ShowDialog()