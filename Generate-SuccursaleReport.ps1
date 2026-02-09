#Requires -Version 5.1
<#
.SYNOPSIS
    Génère un rapport HTML des employés classés par succursale
.DESCRIPTION
    Lit les données AD et classe les employés par succursale en utilisant
    un matching intelligent des adresses. Génère un rapport HTML formaté.
#>

[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

# ============================================
# CONFIGURATION
# ============================================

$OUPath = "OU=Users,OU=DFL-MTL,OU=Divisions,DC=groupedeschenes,DC=loc"
$SuccursalesFile = "Succursales addresses.xlsx"

# ============================================
# FONCTIONS
# ============================================

function Load-SuccursalesData {
    param([string]$FilePath)

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($FilePath)
        $worksheet = $workbook.Worksheets.Item(1)

        $lastRow = $worksheet.UsedRange.Rows.Count
        $succursales = @()

        # Lire à partir de la ligne 2 (skip header)
        for ($i = 2; $i -le $lastRow; $i++) {
            $nom = $worksheet.Cells.Item($i, 1).Text.Trim()
            $adresse = $worksheet.Cells.Item($i, 2).Text.Trim()
            $numero = $worksheet.Cells.Item($i, 3).Text.Trim()

            if ($nom -and $nom -notlike "*Liste*") {
                $isEspacePlomberium = $numero -in @('21', '23', '24', '25', '26', '27', '50')

                $succursales += [PSCustomObject]@{
                    Numero = $numero
                    Nom = $nom
                    Adresse = $adresse
                    Type = if ($isEspacePlomberium) { "Espace Plombérium" } else { "Succursale" }
                    MotsCles = (Get-AddressKeywords $adresse)
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

        return $succursales
    }
    catch {
        Write-Error "Erreur lecture fichier succursales: $($_.Exception.Message)"
        return $null
    }
}

function Get-AddressKeywords {
    param([string]$Address)

    if ([string]::IsNullOrWhiteSpace($Address)) {
        return @()
    }

    # Extraire mots-clés: numéros de rue, noms de rues, villes
    $keywords = @()

    # Normaliser l'adresse
    $normalized = $Address.ToLower() -replace '[,\.]', ' '

    # Extraire numéros (de rue)
    $numbers = [regex]::Matches($normalized, '\b\d{2,5}\b')
    foreach ($num in $numbers) {
        $keywords += $num.Value
    }

    # Mots importants (noms de rues, villes)
    $words = $normalized -split '\s+'
    foreach ($word in $words) {
        if ($word.Length -ge 3 -and $word -notmatch '^\d+$') {
            # Exclure mots communs
            if ($word -notin @('rue', 'boul', 'boulevard', 'avenue', 'ave', 'chemin', 'ch', 'montée', 'autoroute', 'voie')) {
                $keywords += $word
            }
        }
    }

    return $keywords | Select-Object -Unique
}

function Match-AddressToSuccursale {
    param(
        [string]$UserAddress,
        [string]$UserCity,
        [string]$UserPostalCode,
        $Succursales
    )

    if ([string]::IsNullOrWhiteSpace($UserAddress)) {
        return $null
    }

    $userKeywords = Get-AddressKeywords "$UserAddress $UserCity"
    $bestMatch = $null
    $bestScore = 0

    foreach ($succ in $Succursales) {
        $score = 0

        # Comparer mots-clés
        foreach ($keyword in $userKeywords) {
            if ($succ.MotsCles -contains $keyword) {
                $score += 10
            }
            # Fuzzy match partiel
            foreach ($succKey in $succ.MotsCles) {
                if ($succKey.Contains($keyword) -or $keyword.Contains($succKey)) {
                    $score += 5
                }
            }
        }

        # Bonus si le nom de la succursale est dans l'adresse ou la ville
        $succNom = $succ.Nom.ToLower() -replace '\s+', ''
        $userText = "$UserAddress $UserCity".ToLower() -replace '\s+', ''
        if ($userText.Contains($succNom) -or $succNom.Contains($userText.Substring(0, [Math]::Min(8, $userText.Length)))) {
            $score += 20
        }

        if ($score -gt $bestScore) {
            $bestScore = $score
            $bestMatch = $succ
        }
    }

    # Retourner seulement si score suffisant
    if ($bestScore -ge 10) {
        return $bestMatch
    }

    return $null
}

function Get-AllPhoneDirectory {
    param([string]$OUSearchBase)

    try {
        Write-Host "Chargement des utilisateurs depuis AD..."

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

            # Inclure si extension OU adresse présente
            if ([string]::IsNullOrWhiteSpace($extension) -and [string]::IsNullOrWhiteSpace($address)) {
                continue
            }

            $city = $user.City
            if (-not $city) { $city = $user.l }
            if (-not $city) { $city = $user.Office }

            $postalCode = $user.PostalCode

            $allUsers += [PSCustomObject]@{
                Nom = if ($user.Surname) { $user.Surname } else { "" }
                Prenom = if ($user.GivenName) { $user.GivenName } else { "" }
                Adresse = if ($address) { $address } else { "" }
                Ville = if ($city) { $city } else { "" }
                CodePostal = if ($postalCode) { $postalCode } else { "" }
                Extension = if ($extension) { $extension } else { "" }
                Email = if ($user.EmailAddress) { $user.EmailAddress } else { "" }
                Titre = if ($user.Title) { $user.Title } else { "" }
                Departement = if ($user.Department) { $user.Department } else { "" }
                SamAccountName = $user.SamAccountName
            }
        }

        Write-Host "Total utilisateurs chargés: $($allUsers.Count)"
        return $allUsers | Sort-Object Nom, Prenom
    }
    catch {
        Write-Error "Erreur AD: $($_.Exception.Message)"
        return $null
    }
}

function Generate-HTMLReport {
    param(
        $Succursales,
        $EmployesBySuccursale,
        $UnclassifiedUsers,
        [string]$OutputPath
    )

    $html = @"
<!DOCTYPE html>
<html lang="fr">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Répertoire Téléphonique par Succursale</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            color: #333;
        }
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
            overflow: hidden;
        }
        .header {
            background: linear-gradient(135deg, #0078d7 0%, #005a9e 100%);
            color: white;
            padding: 40px;
            text-align: center;
        }
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            text-shadow: 2px 2px 4px rgba(0,0,0,0.3);
        }
        .header .subtitle {
            font-size: 1.2em;
            opacity: 0.9;
        }
        .stats {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            padding: 30px;
            background: #f8f9fa;
        }
        .stat-box {
            background: white;
            padding: 20px;
            border-radius: 10px;
            text-align: center;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
        }
        .stat-number {
            font-size: 2.5em;
            font-weight: bold;
            color: #0078d7;
        }
        .stat-label {
            color: #666;
            margin-top: 5px;
        }
        .content {
            padding: 30px;
        }
        .toc {
            background: #f0f8ff;
            padding: 25px;
            border-radius: 10px;
            margin-bottom: 30px;
            border-left: 5px solid #0078d7;
        }
        .toc h2 {
            color: #0078d7;
            margin-bottom: 15px;
        }
        .toc ul {
            list-style: none;
        }
        .toc li {
            padding: 8px 0;
        }
        .toc a {
            color: #0078d7;
            text-decoration: none;
            font-weight: 500;
            transition: all 0.3s;
        }
        .toc a:hover {
            color: #005a9e;
            padding-left: 10px;
        }
        .succursale-section {
            margin: 40px 0;
            page-break-inside: avoid;
        }
        .succursale-header {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
            padding: 25px;
            border-radius: 10px;
            margin-bottom: 20px;
            box-shadow: 0 4px 15px rgba(0,0,0,0.2);
        }
        .succursale-header.espace-plomberium {
            background: linear-gradient(135deg, #f093fb 0%, #f5576c 100%);
        }
        .succursale-header h2 {
            font-size: 1.8em;
            margin-bottom: 10px;
        }
        .succursale-info {
            opacity: 0.95;
            margin-top: 10px;
        }
        .employee-table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 20px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            border-radius: 10px;
            overflow: hidden;
        }
        .employee-table thead {
            background: #0078d7;
            color: white;
        }
        .employee-table th {
            padding: 15px;
            text-align: left;
            font-weight: 600;
        }
        .employee-table tbody tr {
            border-bottom: 1px solid #eee;
            transition: background 0.3s;
        }
        .employee-table tbody tr:hover {
            background: #f8f9fa;
        }
        .employee-table tbody tr:nth-child(even) {
            background: #fafafa;
        }
        .employee-table tbody tr:nth-child(even):hover {
            background: #f0f0f0;
        }
        .employee-table td {
            padding: 12px 15px;
        }
        .badge {
            display: inline-block;
            padding: 4px 12px;
            border-radius: 20px;
            font-size: 0.85em;
            font-weight: 600;
        }
        .badge-primary {
            background: #e3f2fd;
            color: #1976d2;
        }
        .badge-success {
            background: #e8f5e9;
            color: #388e3c;
        }
        .no-employees {
            text-align: center;
            padding: 40px;
            color: #999;
            font-style: italic;
        }
        .footer {
            background: #f8f9fa;
            padding: 30px;
            text-align: center;
            color: #666;
            border-top: 3px solid #0078d7;
        }
        @media print {
            body { background: white; padding: 0; }
            .container { box-shadow: none; }
            .succursale-section { page-break-inside: avoid; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>📞 RÉPERTOIRE TÉLÉPHONIQUE</h1>
            <div class="subtitle">Classification par Succursale et Espace Plombérium</div>
            <div class="subtitle" style="margin-top: 10px; font-size: 0.9em;">
                Généré le $(Get-Date -Format 'dd MMMM yyyy à HH:mm')
            </div>
        </div>
"@

    # Statistiques
    $totalEmployees = ($EmployesBySuccursale.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum
    $totalSuccursales = $Succursales | Where-Object { $_.Type -eq "Succursale" } | Measure-Object | Select-Object -ExpandProperty Count
    $totalEspaces = $Succursales | Where-Object { $_.Type -eq "Espace Plombérium" } | Measure-Object | Select-Object -ExpandProperty Count

    $html += @"
        <div class="stats">
            <div class="stat-box">
                <div class="stat-number">$totalEmployees</div>
                <div class="stat-label">Employés Classés</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$totalSuccursales</div>
                <div class="stat-label">Succursales</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$totalEspaces</div>
                <div class="stat-label">Espaces Plombérium</div>
            </div>
            <div class="stat-box">
                <div class="stat-number">$($UnclassifiedUsers.Count)</div>
                <div class="stat-label">Non Classés</div>
            </div>
        </div>

        <div class="content">
"@

    # Table des matières
    $html += @"
            <div class="toc">
                <h2>📋 Table des Matières</h2>
                <ul>
"@

    # TOC Succursales
    $html += "<li><strong>Succursales:</strong><ul>"
    foreach ($succ in ($Succursales | Where-Object { $_.Type -eq "Succursale" } | Sort-Object Numero)) {
        $count = if ($EmployesBySuccursale.ContainsKey($succ.Numero)) { $EmployesBySuccursale[$succ.Numero].Count } else { 0 }
        $html += "<li><a href='#succ$($succ.Numero)'>$($succ.Nom) (#$($succ.Numero)) - $count employés</a></li>"
    }
    $html += "</ul></li>"

    # TOC Espaces
    $html += "<li><strong>Espaces Plombérium:</strong><ul>"
    foreach ($succ in ($Succursales | Where-Object { $_.Type -eq "Espace Plombérium" } | Sort-Object Numero)) {
        $count = if ($EmployesBySuccursale.ContainsKey($succ.Numero)) { $EmployesBySuccursale[$succ.Numero].Count } else { 0 }
        $html += "<li><a href='#succ$($succ.Numero)'>$($succ.Nom) (#$($succ.Numero)) - $count employés</a></li>"
    }
    $html += "</ul></li>"

    if ($UnclassifiedUsers.Count -gt 0) {
        $html += "<li><a href='#unclassified'>Non Classés - $($UnclassifiedUsers.Count) employés</a></li>"
    }

    $html += @"
                </ul>
            </div>
"@

    # Sections par succursale
    foreach ($succ in ($Succursales | Sort-Object { if ($_.Type -eq "Succursale") { 0 } else { 1 } }, Numero)) {
        $employees = if ($EmployesBySuccursale.ContainsKey($succ.Numero)) { $EmployesBySuccursale[$succ.Numero] } else { @() }
        $cssClass = if ($succ.Type -eq "Espace Plombérium") { "espace-plomberium" } else { "" }

        $html += @"
            <div class="succursale-section" id="succ$($succ.Numero)">
                <div class="succursale-header $cssClass">
                    <h2>$($succ.Nom) <span class="badge badge-primary">#$($succ.Numero)</span></h2>
                    <div class="succursale-info">
                        <strong>📍 Adresse:</strong> $($succ.Adresse)<br>
                        <strong>🏢 Type:</strong> $($succ.Type)<br>
                        <strong>👥 Employés:</strong> $($employees.Count)
                    </div>
                </div>
"@

        if ($employees.Count -gt 0) {
            $html += @"
                <table class="employee-table">
                    <thead>
                        <tr>
                            <th>Nom</th>
                            <th>Prénom</th>
                            <th>Extension</th>
                            <th>Email</th>
                            <th>Titre</th>
                            <th>Adresse</th>
                        </tr>
                    </thead>
                    <tbody>
"@
            foreach ($emp in ($employees | Sort-Object Nom, Prenom)) {
                $html += @"
                        <tr>
                            <td><strong>$($emp.Nom)</strong></td>
                            <td>$($emp.Prenom)</td>
                            <td><span class="badge badge-success">$($emp.Extension)</span></td>
                            <td>$($emp.Email)</td>
                            <td>$($emp.Titre)</td>
                            <td style="font-size: 0.9em;">$($emp.Adresse), $($emp.Ville)</td>
                        </tr>
"@
            }
            $html += @"
                    </tbody>
                </table>
"@
        } else {
            $html += @"
                <div class="no-employees">Aucun employé classé dans cette succursale</div>
"@
        }

        $html += "</div>"
    }

    # Non classés
    if ($UnclassifiedUsers.Count -gt 0) {
        $html += @"
            <div class="succursale-section" id="unclassified">
                <div class="succursale-header" style="background: linear-gradient(135deg, #ff9a9e 0%, #fecfef 100%);">
                    <h2>❓ Employés Non Classés</h2>
                    <div class="succursale-info">
                        <strong>👥 Total:</strong> $($UnclassifiedUsers.Count) employés sans succursale identifiée
                    </div>
                </div>
                <table class="employee-table">
                    <thead>
                        <tr>
                            <th>Nom</th>
                            <th>Prénom</th>
                            <th>Extension</th>
                            <th>Email</th>
                            <th>Adresse</th>
                            <th>Ville</th>
                        </tr>
                    </thead>
                    <tbody>
"@
        foreach ($emp in ($UnclassifiedUsers | Sort-Object Nom, Prenom)) {
            $html += @"
                        <tr>
                            <td><strong>$($emp.Nom)</strong></td>
                            <td>$($emp.Prenom)</td>
                            <td><span class="badge badge-success">$($emp.Extension)</span></td>
                            <td>$($emp.Email)</td>
                            <td style="font-size: 0.9em;">$($emp.Adresse)</td>
                            <td>$($emp.Ville)</td>
                        </tr>
"@
        }
        $html += @"
                    </tbody>
                </table>
            </div>
"@
    }

    $html += @"
        </div>

        <div class="footer">
            <p><strong>Groupe Deschênes</strong> - Répertoire Téléphonique par Succursale</p>
            <p style="margin-top: 10px; font-size: 0.9em;">
                Rapport généré automatiquement le $(Get-Date -Format 'dd/MM/yyyy à HH:mm:ss')
            </p>
        </div>
    </div>
</body>
</html>
"@

    $html | Out-File -FilePath $OutputPath -Encoding UTF8
    Write-Host "Rapport généré: $OutputPath" -ForegroundColor Green
}

# ============================================
# MAIN
# ============================================

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "  GÉNÉRATEUR DE RAPPORT PAR SUCCURSALE" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
Write-Host ""

# 1. Charger succursales
Write-Host "[1/4] Chargement des succursales..." -ForegroundColor Yellow
$succursales = Load-SuccursalesData -FilePath (Join-Path $PSScriptRoot $SuccursalesFile)
if (-not $succursales) {
    Write-Error "Impossible de charger les succursales"
    exit 1
}
Write-Host "✓ $($succursales.Count) succursales chargées" -ForegroundColor Green
Write-Host ""

# 2. Charger employés AD
Write-Host "[2/4] Chargement des employés depuis AD..." -ForegroundColor Yellow
$employees = Get-AllPhoneDirectory -OUSearchBase $OUPath
if (-not $employees) {
    Write-Error "Impossible de charger les employés"
    exit 1
}
Write-Host "✓ $($employees.Count) employés chargés" -ForegroundColor Green
Write-Host ""

# 3. Classifier les employés
Write-Host "[3/4] Classification des employés par succursale..." -ForegroundColor Yellow
$employeesBySuccursale = @{}
$unclassifiedUsers = @()

foreach ($emp in $employees) {
    $match = Match-AddressToSuccursale -UserAddress $emp.Adresse -UserCity $emp.Ville -UserPostalCode $emp.CodePostal -Succursales $succursales

    if ($match) {
        if (-not $employeesBySuccursale.ContainsKey($match.Numero)) {
            $employeesBySuccursale[$match.Numero] = @()
        }
        $employeesBySuccursale[$match.Numero] += $emp
    } else {
        $unclassifiedUsers += $emp
    }
}

Write-Host "✓ Classification terminée" -ForegroundColor Green
Write-Host "  - Employés classés: $(($employeesBySuccursale.Values | ForEach-Object { $_.Count } | Measure-Object -Sum).Sum)" -ForegroundColor Cyan
Write-Host "  - Non classés: $($unclassifiedUsers.Count)" -ForegroundColor Cyan
Write-Host ""

# 4. Générer rapport HTML
Write-Host "[4/4] Génération du rapport HTML..." -ForegroundColor Yellow
$outputPath = Join-Path $PSScriptRoot "Rapport_Succursales_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
Generate-HTMLReport -Succursales $succursales -EmployesBySuccursale $employeesBySuccursale -UnclassifiedUsers $unclassifiedUsers -OutputPath $outputPath

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  RAPPORT GÉNÉRÉ AVEC SUCCÈS !" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Fichier: $outputPath" -ForegroundColor Cyan
Write-Host ""
Write-Host "Ouvrir le rapport? (O/N)" -ForegroundColor Yellow -NoNewline
$response = Read-Host " "
if ($response -eq 'O' -or $response -eq 'o') {
    Start-Process $outputPath
}
