
# Imposta la cartella sorgente (share con sottocartelle)
$sourceFolder = "\\SWOL003D\GBilder\03_TEXTE\Cumulus"

# Imposta la cartella di destinazione
$destinationFolder = "\\SWOL003D\GBilder\03_TEXTE\Cumulus\ExcelFiles"

# Crea la cartella di destinazione se non esiste
if (-not (Test-Path -Path $destinationFolder)) {
    New-Item -ItemType Directory -Path $destinationFolder
}

# Trova tutti i file .docx ricorsivamente
Get-ChildItem -Path $sourceFolder -Recurse -Filter *.xlsx | ForEach-Object {
    $sourceFile = $_.FullName
    $destinationFile = Join-Path -Path $destinationFolder -ChildPath $_.Name
    # Se esiste già un file con lo stesso nome, aggiunge un numero per evitare conflitti
    $counter = 1
    while (Test-Path $destinationFile) {
        $baseName = [System.IO.Path]::GetFileNameWithoutExtension($_.Name)
        $extension = $_.Extension
        $destinationFile = Join-Path -Path $destinationFolder -ChildPath \"${baseName}_$counter$extension\"
        $counter++
    }

    Copy-Item -Path $sourceFile -Destination $destinationFile
}
