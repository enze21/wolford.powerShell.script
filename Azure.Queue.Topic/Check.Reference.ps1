param(
    [string]$LibPath = "D:\PowerShell\Azure.Queue.Topic\libs"
)

Write-Host "Controllo DLL nella cartella: $LibPath" -ForegroundColor Cyan

# DLL presenti nella cartella
$existingDlls = Get-ChildItem -Path $LibPath -Filter *.dll | Select-Object -ExpandProperty Name

# Lista risultati
$results = @()

foreach ($dll in Get-ChildItem -Path $LibPath -Filter *.dll) {

    try {
        $assembly = [System.Reflection.Assembly]::ReflectionOnlyLoadFrom($dll.FullName)

        foreach ($ref in $assembly.GetReferencedAssemblies()) {

            $expectedName = "$($ref.Name).dll"
            $present = $existingDlls -contains $expectedName

            $results += [PSCustomObject]@{
                DLL               = $dll.Name
                Dependency        = $ref.Name
                Version           = $ref.Version
                ExpectedFile      = $expectedName
                PresentInLibs     = $present
            }
        }
    }
    catch {
        $results += [PSCustomObject]@{
            DLL               = $dll.Name
            Dependency        = "[ERROR READING DLL]"
            Version           = ""
            ExpectedFile      = ""
            PresentInLibs     = $false
        }
    }
}

# Filtra system base (non da scaricare)
$ignore = @('mscorlib', 'System', 'System.Core', 'netstandard')

$missing = $results | Where-Object {
    $_.PresentInLibs -eq $false -and
    $_.Dependency -notin $ignore -and
    $_.Dependency -ne "[ERROR READING DLL]"
}

Write-Host ""
Write-Host "===== DLL MANCANTI =====" -ForegroundColor Yellow

if ($missing.Count -eq 0) {
    Write-Host "Nessuna dipendenza mancante ✅" -ForegroundColor Green
}
else {
    $missing |
        Sort-Object Dependency, Version |
        Format-Table Dependency, Version, ExpectedFile -AutoSize
}

# Salva anche output su file (utile)
$outputFile = Join-Path $LibPath "missing-dependencies.txt"
$missing | Sort-Object Dependency, Version | Out-File $outputFile

Write-Host ""
Write-Host "Report salvato in: $outputFile" -ForegroundColor Cyan