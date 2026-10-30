#Requires -Version 7.2
#Requires -Modules Az.Accounts, Az.Storage

[CmdletBinding()]
param(
    [string]$SubscriptionId = "a0b81247-87e8-4e11-880e-83c23340a726",
    [string]$StorageAccountName = "stowofosftpdev",
    [string]$ContainerName = "stealthftp",
    [string]$DestinationPath = "D:\StealthFTP\Inbound",
    [string]$StagingPath = "D:\StealthFTP\Staging",
    [string]$LogDirectory = "D:\StealthFTP\Logs",
    [string]$BlobPrefix = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

$script:ExitCode = 0
$script:DownloadedCount = 0
$script:DeletedCount = 0
$script:PublishedCount = 0
$script:SkippedCount = 0
$script:FailedCount = 0
$script:TotalBytesDownloaded = [int64]0
$script:StartTime = Get-Date
$script:LogFile = $null
$script:Mutex = $null
$script:HasMutex = $false

function New-DirectoryIfMissing {
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"
    $line = "$timestamp [$Level] $Message"

    switch ($Level) {
        "ERROR"   { Write-Host $line -ForegroundColor Red }
        "WARN"    { Write-Host $line -ForegroundColor Yellow }
        "SUCCESS" { Write-Host $line -ForegroundColor Green }
        default   { Write-Host $line }
    }

    if ($script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $line -Encoding utf8
    }
}

function Get-SafeLocalPath {
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$BlobName
    )

    $relativePath = $BlobName.Replace('/', [IO.Path]::DirectorySeparatorChar)
    $fullBasePath = [IO.Path]::GetFullPath($BasePath)
    $fullLocalPath = [IO.Path]::GetFullPath((Join-Path -Path $fullBasePath -ChildPath $relativePath))
    $baseWithSeparator = $fullBasePath.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar

    if (-not $fullLocalPath.StartsWith($baseWithSeparator, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Percorso locale non valido per il blob '$BlobName'."
    }

    return $fullLocalPath
}

function Publish-StagedFile {
    param(
        [Parameter(Mandatory)][string]$StagedFile,
        [Parameter(Mandatory)][string]$FinalFile
    )

    $finalDirectory = Split-Path -Path $FinalFile -Parent
    New-DirectoryIfMissing -Path $finalDirectory

    if (Test-Path -LiteralPath $FinalFile -PathType Leaf) {
        $backupFile = "$FinalFile.replaced.$((Get-Date).ToString('yyyyMMddHHmmssfff'))"
        Move-Item -LiteralPath $FinalFile -Destination $backupFile -Force
        Write-Log -Level "WARN" -Message "File locale esistente spostato in backup: $backupFile"
    }

    Move-Item -LiteralPath $StagedFile -Destination $FinalFile -Force
}

try {
    New-DirectoryIfMissing -Path $DestinationPath
    New-DirectoryIfMissing -Path $StagingPath
    New-DirectoryIfMissing -Path $LogDirectory

    $script:LogFile = Join-Path -Path $LogDirectory -ChildPath ("Download-StealthFTP_{0}.log" -f (Get-Date -Format "yyyyMMdd"))

    $script:Mutex = [Threading.Mutex]::new($false, "Global\Download-StealthFTP_$StorageAccountName`_$ContainerName")
    $script:HasMutex = $script:Mutex.WaitOne([TimeSpan]::FromSeconds(5))
    if (-not $script:HasMutex) {
        Write-Log -Level "WARN" -Message "Un'altra istanza dello script e' gia in esecuzione."
        $script:ExitCode = 2
        throw "Esecuzione simultanea non consentita."
    }

    Write-Log -Message "============================================================"
    Write-Log -Message "Avvio trasferimento Azure Blob verso cartella on-premise."
    Write-Log -Message "Storage account: $StorageAccountName"
    Write-Log -Message "Container: $ContainerName"
    Write-Log -Message "Prefix: '$BlobPrefix'"
    Write-Log -Message "Staging: $StagingPath"
    Write-Log -Message "Destinazione: $DestinationPath"

    Disable-AzContextAutosave -Scope Process | Out-Null

    Write-Log -Message "Autenticazione tramite Managed Identity del server Azure Arc."
    Connect-AzAccount -Identity -Subscription $SubscriptionId -Scope Process -ErrorAction Stop | Out-Null
    Set-AzContext -SubscriptionId $SubscriptionId -Scope Process -ErrorAction Stop | Out-Null
    Write-Log -Level "SUCCESS" -Message "Autenticazione Azure completata."

    $storageContext = New-AzStorageContext -StorageAccountName $StorageAccountName -UseConnectedAccount -ErrorAction Stop

    if ([string]::IsNullOrWhiteSpace($BlobPrefix)) {
        $blobs = @(Get-AzStorageBlob -Container $ContainerName -Context $storageContext -ErrorAction Stop)
    }
    else {
        $blobs = @(Get-AzStorageBlob -Container $ContainerName -Prefix $BlobPrefix -Context $storageContext -ErrorAction Stop)
    }

    Write-Log -Message "Blob individuati: $($blobs.Count)."

    foreach ($blob in $blobs) {
        $blobName = $blob.Name
        $temporaryFile = $null
        $stagedFile = $null
        $finalFile = $null
        $blobDeleted = $false

        try {
            if ($blobName.EndsWith('/')) {
                Write-Log -Message "Directory virtuale ignorata: $blobName"
                $script:SkippedCount++
                continue
            }

            $stagedFile = Get-SafeLocalPath -BasePath $StagingPath -BlobName $blobName
            $finalFile = Get-SafeLocalPath -BasePath $DestinationPath -BlobName $blobName
            $temporaryFile = "$stagedFile.downloading"

            New-DirectoryIfMissing -Path (Split-Path -Path $stagedFile -Parent)

            if (Test-Path -LiteralPath $temporaryFile -PathType Leaf) {
                Remove-Item -LiteralPath $temporaryFile -Force
            }

            if (Test-Path -LiteralPath $stagedFile -PathType Leaf) {
                Remove-Item -LiteralPath $stagedFile -Force
            }

            Write-Log -Message "Download in corso: $blobName"
            Get-AzStorageBlobContent -Container $ContainerName -Blob $blobName -Destination $temporaryFile -Context $storageContext -Force -ErrorAction Stop | Out-Null

            if (-not (Test-Path -LiteralPath $temporaryFile -PathType Leaf)) {
                throw "Il file temporaneo non e' stato creato."
            }

            $downloadedItem = Get-Item -LiteralPath $temporaryFile
            $downloadedSize = [int64]$downloadedItem.Length
            $remoteSize = [int64]$blob.Length

            if ($downloadedSize -ne $remoteSize) {
                throw "Verifica dimensione fallita. Azure: $remoteSize byte; locale: $downloadedSize byte."
            }

            Move-Item -LiteralPath $temporaryFile -Destination $stagedFile -Force
            $script:DownloadedCount++
            $script:TotalBytesDownloaded += $downloadedSize
            Write-Log -Level "SUCCESS" -Message "Download verificato in staging: $blobName; $downloadedSize byte."

            Write-Log -Message "Cancellazione blob da Azure: $blobName"
            Remove-AzStorageBlob -Container $ContainerName -Blob $blobName -Context $storageContext -Force -ErrorAction Stop | Out-Null
            $blobDeleted = $true
            $script:DeletedCount++
            Write-Log -Level "SUCCESS" -Message "Blob eliminato da Azure: $blobName"

            Publish-StagedFile -StagedFile $stagedFile -FinalFile $finalFile
            $script:PublishedCount++
            Write-Log -Level "SUCCESS" -Message "File pubblicato nella cartella finale: $finalFile"
        }
        catch {
            $script:FailedCount++
            $script:ExitCode = 1
            Write-Log -Level "ERROR" -Message "Errore sul blob '$blobName': $($_.Exception.Message)"

            if ($temporaryFile -and (Test-Path -LiteralPath $temporaryFile -PathType Leaf)) {
                Remove-Item -LiteralPath $temporaryFile -Force -ErrorAction SilentlyContinue
                Write-Log -Level "WARN" -Message "File temporaneo incompleto rimosso: $temporaryFile"
            }

            if ($blobDeleted -and $stagedFile -and (Test-Path -LiteralPath $stagedFile -PathType Leaf)) {
                Write-Log -Level "ERROR" -Message "Il blob e' stato eliminato ma il file non e' stato pubblicato. Il file integro resta in staging: $stagedFile"
            }
            elseif (-not $blobDeleted) {
                Write-Log -Level "WARN" -Message "Il blob non e' stato eliminato e potra essere ritentato alla prossima esecuzione."
            }
        }
    }
}
catch {
    if ($script:ExitCode -eq 0) {
        $script:ExitCode = 1
    }
    Write-Log -Level "ERROR" -Message "Errore generale: $($_.Exception.Message)"
}
finally {
    $duration = (Get-Date) - $script:StartTime
    $downloadedMB = [Math]::Round($script:TotalBytesDownloaded / 1MB, 2)

    Write-Log -Message "File scaricati e verificati: $script:DownloadedCount"
    Write-Log -Message "Blob eliminati da Azure: $script:DeletedCount"
    Write-Log -Message "File pubblicati localmente: $script:PublishedCount"
    Write-Log -Message "Elementi ignorati: $script:SkippedCount"
    Write-Log -Message "Elementi in errore: $script:FailedCount"
    Write-Log -Message "Totale scaricato: $downloadedMB MB"
    Write-Log -Message ("Durata: {0:hh\:mm\:ss}" -f $duration)
    Write-Log -Message "Exit code: $script:ExitCode"
    Write-Log -Message "============================================================"

    Disconnect-AzAccount -Scope Process -ErrorAction SilentlyContinue | Out-Null

    if ($script:HasMutex -and $script:Mutex) {
        $script:Mutex.ReleaseMutex()
    }
    if ($script:Mutex) {
        $script:Mutex.Dispose()
    }
}

exit $script:ExitCode
