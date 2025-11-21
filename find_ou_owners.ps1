Import-Module ActiveDirectory

$baseOU = ""

function Get-GroupManager {
    param (
        [string]$distinguishedName
    )
    try {
        $group = Get-ADGroup -Identity $distinguishedName -Properties ManagedBy, Description
        if ($group.ManagedBy) {
            try {
                $manager = Get-ADObject -Identity $group.ManagedBy -Properties Name, DisplayName, SamAccountName, ObjectClass
                if ($manager.ObjectClass -eq "user") {
                    $userDetails = Get-ADUser -Identity $manager.DistinguishedName -Properties EmailAddress
                    $email = if ($userDetails.EmailAddress) { " - $($userDetails.EmailAddress)" } else { "" }
                    return [PSCustomObject]@{
                        Nom = $manager.Name
                        Login = $manager.SamAccountName
                        Email = $userDetails.EmailAddress
                        Type = "Utilisateur"
                        CompleteName = "$($manager.Name) ($($manager.SamAccountName))$email"
                    }
                }
                else {
                    return [PSCustomObject]@{
                        Nom = $manager.Name
                        Login = $manager.SamAccountName
                        Email = ""
                        Type = "Groupe"
                        CompleteName = "$($manager.Name) ($($manager.SamAccountName)) [Groupe]"
                    }
                }
            }
            catch {
                return [PSCustomObject]@{
                    Nom = $group.ManagedBy
                    Login = ""
                    Email = ""
                    Type = "Non resolu"
                    CompleteName = $group.ManagedBy
                }
            }
        }
        else {
            return [PSCustomObject]@{
                Nom = ""
                Login = ""
                Email = ""
                Type = "Aucun"
                CompleteName = "Aucun manageur defini"
            }
        }
    }
    catch {
        return [PSCustomObject]@{
            Nom = ""
            Login = ""
            Email = ""
            Type = "Erreur"
            CompleteName = "Erreur: $($_.Exception.Message)"
        }
    }
}

Write-Host "operation 1 demarrage"

try {
    $ouExists = Get-ADOrganizationalUnit -Filter "DistinguishedName -eq '$baseOU'" -ErrorAction SilentlyContinue
    if (-not $ouExists) {
        Write-Host "fail ou inexistante"
        exit
    }

    $allGroups = Get-ADGroup -Filter * -SearchBase $baseOU -SearchScope Subtree -Properties Description, ManagedBy, MemberOf
    Write-Host "groupes trouves: $($allGroups.Count)"

    $results = @()
    $counter = 0

    foreach ($group in $allGroups) {
        $counter++
        Write-Progress -Activity "analyse" -Status "groupe $counter sur $($allGroups.Count)" -PercentComplete (($counter / $allGroups.Count) * 100)
        $manager = Get-GroupManager -distinguishedName $group.DistinguishedName

        $memberCount = ($group.MemberOf | Measure-Object).Count
        
        $results += [PSCustomObject]@{
            'Nom du Groupe' = $group.Name
            'Manageur' = if ($manager.Nom) { $manager.Nom } else { "Aucun" }
        }
    }

    Write-Progress -Activity "analyse" -Completed
    Write-Host "operation 2 analyse ok"

    $exportPath = "$env:USERPROFILE\Desktop\output.xlsx"
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    $workbook = $excel.Workbooks.Add()
    $worksheet = $workbook.Worksheets.Item(1)
    $worksheet.Name = "Groupes et Manageurs"

    $worksheet.Cells.Item(1, 1) = "Nom du Groupe"
    $worksheet.Cells.Item(1, 2) = "Manageur"

    $headerRange = $worksheet.Range("A1:B1")
    $headerRange.Font.Bold = $true
    $headerRange.Font.Size = 12

    $row = 2
    foreach ($result in $results) {
        $worksheet.Cells.Item($row, 1) = $result.'Nom du Groupe'
        $worksheet.Cells.Item($row, 2) = $result.Manageur
        $row++
    }

    $worksheet.Columns.Item(1).AutoFit() | Out-Null
    $worksheet.Columns.Item(2).AutoFit() | Out-Null

    $dataRange = $worksheet.Range("A1:B$($row - 1)")
    $dataRange.Borders.LineStyle = 1
    $dataRange.Borders.Weight = 2

    $worksheet.Range("A1:B1").AutoFilter() | Out-Null
    $worksheet.Application.ActiveWindow.SplitRow = 1
    $worksheet.Application.ActiveWindow.FreezePanes = $true

    $workbook.SaveAs($exportPath)
    $workbook.Close()
    $excel.Quit()

    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()

    Write-Host "operation 3 export ok"
}
catch {
    Write-Host "fail: $($_.Exception.Message)"
}

Write-Host "termine"
