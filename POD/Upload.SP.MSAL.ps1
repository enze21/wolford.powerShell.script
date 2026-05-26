
Unblock-File -Path "D:\PowerShell\Microsoft.SharePoint.Client.dll"
Unblock-File -Path "D:\PowerShell\Microsoft.SharePoint.Client.Runtime.dll"
Unblock-File -Path "D:\PowerShell\Microsoft.IdentityModel.Abstractions.dll"


Add-Type -Path "D:\PowerShell\Microsoft.SharePoint.Client.dll"
Add-Type -Path "D:\PowerShell\Microsoft.SharePoint.Client.Runtime.dll"
Add-Type -Path "D:\PowerShell\Microsoft.IdentityModel.Abstractions.dll"

# CONFIGURAZIONE
$clientId = "a2817a49-d823-4367-aa95-e0cce1c39058"
$tenantId = "eef0b564-afde-4087-948d-ff7180bc51b0"
$siteUrl = "https://wolford365.sharepoint.com/sites/WolfordShare"
$targetFolder = "Documenti/Product Images"  # Path relativo su SharePoint
$localFolderPath = "\\SWOL002D\GBilder\eCommerce\eCommerce\DATABASE IMAGES"
$logFile = "H:\UploadLog.SP.REST.txt"

# AUTENTICAZIONE MODERNA (MSAL)
Import-Module Microsoft.Identity.Client
$authority = "https://login.microsoftonline.com/$tenantId"
$scopes = @("$siteUrl/.default")
$msal = [Microsoft.Identity.Client.PublicClientApplicationBuilder]::Create($clientId).WithAuthority($authority).Build()
$tokenResult = $msal.AcquireTokenInteractive($scopes).ExecuteAsync().Result
$accessToken = $tokenResult.AccessToken

# FUNZIONE UPLOAD FILE
function Upload-FileToSharePoint {
    param (
        [string]$filePath,
        [string]$relativeTargetPath
    )
    try {
        $fileName = [System.IO.Path]::GetFileName($filePath)
        $uploadUrl = "$siteUrl/_api/web/GetFolderByServerRelativeUrl('$relativeTargetPath')/Files/add(url='$fileName',overwrite=true)"
        $headers = @{
            "Authorization" = "Bearer $accessToken"
            "Accept" = "application/json;odata=verbose"
        }
        $fileContent = [System.IO.File]::ReadAllBytes($filePath)
        Invoke-RestMethod -Uri $uploadUrl -Method POST -Headers $headers -Body $fileContent
        Add-Content -Path $logFile -Value "SUCCESS: $filePath -> $relativeTargetPath/$fileName"
        Write-Host "✅ SUCCESS: $filePath -> $relativeTargetPath/$fileName"
    }
    catch {
        Add-Content -Path $logFile -Value "ERROR: $filePath -> $relativeTargetPath/$fileName | $_"
        Write-Host "❌ ERROR: $filePath -> $relativeTargetPath/$fileName | $_"
    }
}

# FUNZIONE RICORSIVA UPLOAD CARTELLE
function Upload-FolderRecursively {
    param (
        [string]$localPath,
        [string]$sharePointPath
    )
    # Upload dei file nella cartella corrente
    Get-ChildItem -Path $localPath -File | ForEach-Object {
        Upload-FileToSharePoint -filePath $_.FullName -relativeTargetPath $sharePointPath
    }
    # Ricorsione per sottocartelle
    Get-ChildItem -Path $localPath -Directory | ForEach-Object {
        $subFolderName = $_.Name
        $newSharePointPath = "$sharePointPath/$subFolderName"
        # Crea la cartella su SharePoint (se non esiste)
        $createFolderUrl = "$siteUrl/_api/web/folders"
        $headers = @{
            "Authorization" = "Bearer $accessToken"
            "Accept" = "application/json;odata=verbose"
            "Content-Type" = "application/json;odata=verbose"
        }
        $body = @{
            '__metadata' = @{ 'type' = 'SP.Folder' }
            'ServerRelativeUrl' = $newSharePointPath
        } | ConvertTo-Json -Depth 3
        try {
            Invoke-RestMethod -Uri $createFolderUrl -Method POST -Headers $headers -Body $body
        } catch {}
        # Ricorsione
        Upload-FolderRecursively -localPath $_.FullName -sharePointPath $newSharePointPath
    }
}

# AVVIO UPLOAD
Upload-FolderRecursively -localPath $localFolderPath -sharePointPath $targetFolder
