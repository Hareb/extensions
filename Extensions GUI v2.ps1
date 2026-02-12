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

# Couleurs
$colorPrimary = [System.Drawing.Color]::FromArgb(0, 120, 215)
$colorSecondary = [System.Drawing.Color]::FromArgb(243, 242, 241)
$colorSuccess = [System.Drawing.Color]::FromArgb(16, 124, 16)
$colorWarning = [System.Drawing.Color]::FromArgb(255, 185, 0)
$colorDanger = [System.Drawing.Color]::FromArgb(232, 17, 35)

# ============================================
# FONCTIONS
# ============================================

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

            # Disambiguation: St-Jerome 008 vs Espace Plomberium 025
            # Les deux succursales partagent la meme adresse physique (1075 Grand-Heron)
            # On distingue par le domaine email, avec le manager comme fallback
            $emailAddress = $user.EmailAddress
            if ($city -and $city -match "(?i)j.r.me") {
                if ($emailAddress -and $emailAddress -like "*@espaceplomberium.com") {
                    $branch = "Espace Plomberium St-Jerome"
                } elseif ($emailAddress -and $emailAddress -like "*@deschenes.ca") {
                    $branch = "St-Jerome"
                } elseif ($managerName -eq "Yannick Blanchet") {
                    # Fallback: le manager de la 025 est toujours Yannick Blanchet
                    $branch = "Espace Plomberium St-Jerome"
                }
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

function Compare-Data {
    param(
        $ADData,
        $FileData
    )
    
    if (-not $ADData -or -not $FileData) {
        return $null
    }
    
    $adSams = $ADData.SamAccountName
    $fileSams = $FileData.SamAccountName
    
    $nouveaux = $ADData | Where-Object { $fileSams -notcontains $_.SamAccountName }
    $partis = $FileData | Where-Object { $adSams -notcontains $_.SamAccountName }
    
    return @{
        Nouveaux = $nouveaux
        Partis = $partis
    }
}

# ============================================
# INTERFACE GRAPHIQUE
# ============================================

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Comparateur de Repertoire Telephonique'
$form.Size = New-Object System.Drawing.Size(1400, 800)
$form.StartPosition = 'CenterScreen'
$form.BackColor = $colorSecondary
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)

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
    
    $script:adData = Get-AllPhoneDirectory -OUSearchBase $OUPath -LocationMap $locationMapping
    
    if ($script:adData) {
        $dataGridAD.Rows.Clear()
        foreach ($user in $script:adData) {
            [void]$dataGridAD.Rows.Add(
                $user.Succursale, $user.Nom, $user.Prenom, $user.Extension,
                $user.Ville, $user.Email, $user.SamAccountName
            )
        }
        $lblADCount.Text = "Total: $($script:adData.Count) utilisateurs"
        $lblADCount.ForeColor = $colorSuccess
        
        # Si les deux sont chargés, comparer automatiquement
        if ($script:fileData) {
            Compare-AndDisplay
        }
    }
    
    $btnLoadAD.Enabled = $true
    $btnLoadAD.Text = "CHARGER DEPUIS AD"
})

$lblADCount = New-Object System.Windows.Forms.Label
$lblADCount.Location = New-Object System.Drawing.Point(10, 100)
$lblADCount.Size = New-Object System.Drawing.Size(650, 20)
$lblADCount.Text = 'Aucune donnee chargee'
$lblADCount.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$lblADCount.TextAlign = 'MiddleCenter'

$dataGridAD = New-Object System.Windows.Forms.DataGridView
$dataGridAD.Location = New-Object System.Drawing.Point(10, 130)
$dataGridAD.Size = New-Object System.Drawing.Size(650, 310)
$dataGridAD.AllowUserToAddRows = $false
$dataGridAD.AllowUserToDeleteRows = $false
$dataGridAD.ReadOnly = $true
$dataGridAD.SelectionMode = 'FullRowSelect'
$dataGridAD.AutoSizeColumnsMode = 'Fill'
$dataGridAD.BackgroundColor = [System.Drawing.Color]::White

$cols = @("Succursale", "Nom", "Prenom", "Extension", "Ville", "Email", "SamAccountName")
foreach ($colName in $cols) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.HeaderText = $colName
    $col.Name = $colName
    [void]$dataGridAD.Columns.Add($col)
}

$panelAD.Controls.Add($lblADTitle)
$panelAD.Controls.Add($btnLoadAD)
$panelAD.Controls.Add($lblADCount)
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
        
        $script:fileData = Load-ExcelFile -FilePath $openDialog.FileName
        
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
            
            # Si les deux sont chargés, comparer automatiquement
            if ($script:adData) {
                Compare-AndDisplay
            }
        }
        
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

$dataGridFile = New-Object System.Windows.Forms.DataGridView
$dataGridFile.Location = New-Object System.Drawing.Point(10, 130)
$dataGridFile.Size = New-Object System.Drawing.Size(650, 310)
$dataGridFile.AllowUserToAddRows = $false
$dataGridFile.AllowUserToDeleteRows = $false
$dataGridFile.ReadOnly = $true
$dataGridFile.SelectionMode = 'FullRowSelect'
$dataGridFile.AutoSizeColumnsMode = 'Fill'
$dataGridFile.BackgroundColor = [System.Drawing.Color]::White

foreach ($colName in $cols) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.HeaderText = $colName
    $col.Name = $colName
    [void]$dataGridFile.Columns.Add($col)
}

$panelFile.Controls.Add($lblFileTitle)
$panelFile.Controls.Add($btnLoadFile)
$panelFile.Controls.Add($lblFileCount)
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

$dataGridNouveaux = New-Object System.Windows.Forms.DataGridView
$dataGridNouveaux.Location = New-Object System.Drawing.Point(5, 5)
$dataGridNouveaux.Size = New-Object System.Drawing.Size(1315, 100)
$dataGridNouveaux.AllowUserToAddRows = $false
$dataGridNouveaux.AllowUserToDeleteRows = $false
$dataGridNouveaux.ReadOnly = $true
$dataGridNouveaux.SelectionMode = 'FullRowSelect'
$dataGridNouveaux.AutoSizeColumnsMode = 'Fill'
$dataGridNouveaux.BackgroundColor = [System.Drawing.Color]::FromArgb(220, 255, 220)

foreach ($colName in $cols) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.HeaderText = $colName
    $col.Name = $colName
    [void]$dataGridNouveaux.Columns.Add($col)
}

$tabNouveaux.Controls.Add($dataGridNouveaux)
[void]$tabControlDiff.TabPages.Add($tabNouveaux)

# Onglet Partis
$tabPartis = New-Object System.Windows.Forms.TabPage
$tabPartis.Text = 'EMPLOYES PARTIS'
$tabPartis.BackColor = [System.Drawing.Color]::White

$dataGridPartis = New-Object System.Windows.Forms.DataGridView
$dataGridPartis.Location = New-Object System.Drawing.Point(5, 5)
$dataGridPartis.Size = New-Object System.Drawing.Size(1315, 100)
$dataGridPartis.AllowUserToAddRows = $false
$dataGridPartis.AllowUserToDeleteRows = $false
$dataGridPartis.ReadOnly = $true
$dataGridPartis.SelectionMode = 'FullRowSelect'
$dataGridPartis.AutoSizeColumnsMode = 'Fill'
$dataGridPartis.BackgroundColor = [System.Drawing.Color]::FromArgb(255, 220, 220)

foreach ($colName in $cols) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.HeaderText = $colName
    $col.Name = $colName
    [void]$dataGridPartis.Columns.Add($col)
}

$tabPartis.Controls.Add($dataGridPartis)
[void]$tabControlDiff.TabPages.Add($tabPartis)

$panelDiff.Controls.Add($lblDiffTitle)
$panelDiff.Controls.Add($lblDiffStats)
$panelDiff.Controls.Add($tabControlDiff)
$form.Controls.Add($panelDiff)

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
        
        # Stats
        $nouveauxCount = $comparison.Nouveaux.Count
        $partisCount = $comparison.Partis.Count
        $lblDiffStats.Text = "NOUVEAUX: $nouveauxCount  |  PARTIS: $partisCount"
        
        if ($nouveauxCount -gt 0) {
            $lblDiffStats.ForeColor = $colorSuccess
        } elseif ($partisCount -gt 0) {
            $lblDiffStats.ForeColor = $colorDanger
        } else {
            $lblDiffStats.ForeColor = $colorPrimary
            $lblDiffStats.Text = "Aucune difference detectee"
        }
    }
}

# Afficher le formulaire
[void]$form.ShowDialog()