
# Parametri
$accountName = "stowofodatastorage"
#https://stowofodatastorage.blob.core.windows.net/?sv=2024-11-04&ss=bfqt&srt=sco&sp=rwdlacupiytfx&se=2027-04-01T18:52:55Z&st=2026-02-11T11:37:55Z&spr=https&sig=Phy7Uy%2F4Td0dw4iQmYlV8IG012aeNUoOZ0Qv8dnjos0%3D
$sasToken = "sv=2024-11-04&ss=bfqt&srt=sco&sp=rwdlacupiytfx&se=2027-04-01T18:52:55Z&st=2026-02-11T11:37:55Z&spr=https&sig=Phy7Uy/4Td0dw4iQmYlV8IG012aeNUoOZ0Qv8dnjos0="   # es: sv=2023-11-03&se=2026-02-10...&sig=...


$searchText = "O01K"

$baseUrl = "https://$accountName.blob.core.windows.net"

# === 1. Ottieni lista container, rimuovendo il BOM ===
$containersUrl = "$baseUrl/?comp=list&$sasToken"
$raw = Invoke-WebRequest -Uri $containersUrl -Method GET
$xmlText = $raw.Content -replace "ï»¿", ""   # Rimozione BOM


[xml]$containersXml = $xmlText
$containers = $containersXml.EnumerationResults.Containers.Container.Name

foreach ($c in $containers) {

    Write-Host "Controllo container: $c" -ForegroundColor Cyan

    # === 2. Lista blob nel container ===
    $blobsUrl = "$baseUrl/$c?restype=container&comp=list&$sasToken"
    $rawBlobs = Invoke-WebRequest -Uri $blobsUrl -Method GET
    $blobsXmlText = $rawBlobs.Content -replace "ï»¿", ""

    [xml]$blobsXml = $blobsXmlText
    $blobs = $blobsXml.EnumerationResults.Blobs.Blob.Name

    foreach ($b in $blobs) {

        $blobUrl = "$baseUrl/$c/$b?$sasToken"

        try {
            $content = Invoke-WebRequest -Uri $blobUrl -Method GET -ErrorAction Stop
            $text = $content.Content

            if ($text -match $searchText) {
                Write-Host "🔍 Trovato in: $c/$b" -ForegroundColor Green
            }
        }
        catch {
            Write-Host "⛔ Problema su: $c/$b" -ForegroundColor Red
        }
    }
}
``
