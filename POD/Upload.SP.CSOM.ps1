
# Configurazione
$localFolder = "\\SWOL002D\GBilder\eCommerce\eCommerce\DATABASE IMAGES"
$sharePointUrl = "https://wolford365.sharepoint.com/sites/WolfordShare/Documenti/Product%20Images"
$logFile = "H:\UploadLog.SP.CSOM.txt"

# Credenziali utente
$cred = Get-Credential  # Inserisci le credenziali di Office 365
# Funzione per upload file
function Upload-File {
    param (
        [string]$filePath,
        [string]$targetUrl
    )

    try {
        $webclient = New-Object System.Net.WebClient
        $webclient.Credentials = $cred

        $fileName = [System.IO.Path]::GetFileName($filePath)
        $destination = "$targetUrl/$fileName"

        $webclient.UploadFile($destination, "PUT", $filePath)

        Add-Content -Path $logFile -Value "SUCCESS: $filePath -> $destination"
        Write-Host "✅ SUCCESS: $filePath -> $destination"
    }
    catch {
        Add-Content -Path $logFile -Value "ERROR: $filePath -> $destination | $_"
        Write-Host "❌ ERROR: $filePath -> $destination | $_"

    }
}

# Funzione ricorsiva per upload cartelle
function Upload-Folder {
    param (
        [string]$sourceFolder,
        [string]$targetUrl
    )

    # Upload dei file nella cartella corrente
    Get-ChildItem -Path $sourceFolder -File | ForEach-Object {
        Upload-File -filePath $_.FullName -targetUrl $targetUrl
    }

    # Ricorsione per sottocartelle
    Get-ChildItem -Path $sourceFolder -Directory | ForEach-Object {
        $subFolderName = $_.Name
        $newTargetUrl = "$targetUrl/$subFolderName"

        # SharePoint crea automaticamente le cartelle all'upload del file
        Upload-Folder -sourceFolder $_.FullName -targetUrl $newTargetUrl
    }
}

# Avvio upload
Upload-Folder -sourceFolder $localFolder -targetUrl $sharePointUrl
