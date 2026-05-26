
Install-Module -Name "PnP.PowerShell"

# Parametri
$siteUrl = "https://wolford365.sharepoint.com/sites/WolfordShare"
$localFolderPath = "\\SWOL002D\GBilder\eCommerce\eCommerce\DATABASE IMAGES"
$sharePointRootFolder = "Documenti/Product Images"
$logFilePath = "H:\UploadSPLog.txt"



# Connessione a SharePoint
Connect-PnPOnline -Url $siteUrl -Interactive

# Inizializza il file di log
if (Test-Path $logFilePath) { Remove-Item $logFilePath }
New-Item -Path $logFilePath -ItemType File | Out-Null

# Funzione ricorsiva per upload
function Upload-FolderToSharePoint {
    param (
        [string]$LocalPath,
        [string]$SharePointPath
    )

    # Upload dei file nella cartella corrente
    Get-ChildItem -Path $LocalPath -File | ForEach-Object {
        $targetPath = "$SharePointPath/$($_.Name)"
        try {
            Add-PnPFile -Path $_.FullName -Folder $SharePointPath -ErrorAction Stop
            Add-Content -Path $logFilePath -Value "File caricato: $($_.FullName) → $targetPath"
        } catch {
            Add-Content -Path $logFilePath -Value "❌ Errore nel caricamento file: $($_.FullName) → $targetPath"
        }
    }

    # Ricorsione nelle sottocartelle
    Get-ChildItem -Path $LocalPath -Directory | ForEach-Object {
        $subFolderPath = "$SharePointPath/$($_.Name)"
        try {
            # SharePoint crea automaticamente le cartelle quando si carica un file
            Add-Content -Path $logFilePath -Value "Cartella trovata: $($_.FullName) → $subFolderPath"
        } catch {
            Add-Content -Path $logFilePath -Value "❌ Errore nella creazione cartella: $($_.FullName)"
        }

        Upload-FolderToSharePoint -LocalPath $_.FullName -SharePointPath $subFolderPath
    }
}

# Avvio dell'upload
Upload-FolderToSharePoint -LocalPath $localFolderPath -SharePointPath $sharePointRootFolder
Write-Host "✅ Upload completato. Log disponibile in: $logFilePath"
