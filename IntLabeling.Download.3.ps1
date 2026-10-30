#Requires -Version 5.1
#Requires -Modules Az.Storage

[CmdletBinding()]
param(
    [string]$DestinationPath = "D:\StealthFTP\Inbound",
    [string]$StagingPath     = "D:\StealthFTP\Staging",
    [string]$LogDirectory    = "D:\StealthFTP\Logs",
    [string]$BlobPrefix      = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"

# Azure Storage configuration
$StorageAccountName = "stowofosftpdev"
$ContainerName      = "stealthftp"
# Paste the container SAS here. Required permissions: Read + List + Delete.
$SasToken           = "PASTE_SAS_TOKEN_HERE"

$exitCode = 0
$downloaded = 0
$deleted = 0
$published = 0
$skipped = 0
$failed = 0
$totalBytes = [int64]0
$startTime = Get-Date
$mutex = $null
$hasMutex = $false
$script:LogFile = $null

function Ensure-Directory {
    param([Parameter(Mandatory=$true)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory=$true)][string]$Message,
        [ValidateSet("INFO","WARN","ERROR","SUCCESS")][string]$Level = "INFO"
    )
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"), $Level, $Message
    Write-Host $line
    if ($script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8
    }
}

function Get-SafePath {
    param(
        [Parameter(Mandatory=$true)][string]$BasePath,
        [Parameter(Mandatory=$true)][string]$BlobName
    )
    $base = [IO.Path]::GetFullPath($BasePath)
    $relative = $BlobName.Replace('/', [IO.Path]::DirectorySeparatorChar)
    $result = [IO.Path]::GetFullPath((Join-Path -Path $base -ChildPath $relative))
    $prefix = $base.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
    if (-not $result.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Percorso locale non valido per il blob '$BlobName'."
    }
    return $result
}

try {
    Ensure-Directory -Path $DestinationPath
    Ensure-Directory -Path $StagingPath
    Ensure-Directory -Path $LogDirectory
    $script:LogFile = Join-Path -Path $LogDirectory -ChildPath ("IntLabeling.Download_{0}.log" -f (Get-Date -Format "yyyyMMdd"))

    $mutexName = "Global\IntLabelingDownload_{0}_{1}" -f $StorageAccountName, $ContainerName
    $mutex = New-Object System.Threading.Mutex($false, $mutexName)
    $hasMutex = $mutex.WaitOne([TimeSpan]::FromSeconds(5))
    if (-not $hasMutex) {
        $exitCode = 2
        throw "Un'altra istanza dello script e' gia in esecuzione."
    }

    Write-Log -Message "============================================================"
    Write-Log -Message "Avvio processo. Storage=$StorageAccountName; Container=$ContainerName"
    Write-Log -Message "Staging=$StagingPath"
    Write-Log -Message "Destination=$DestinationPath"

    if ([string]::IsNullOrWhiteSpace($SasToken) -or $SasToken -eq "PASTE_SAS_TOKEN_HERE") {
        throw "SAS non configurata. Inserire il token nella variabile SasToken."
    }

    $effectiveSas = $SasToken.Trim()
    if ($effectiveSas.StartsWith("?")) { $effectiveSas = $effectiveSas.Substring(1) }

    $context = New-AzStorageContext -StorageAccountName $StorageAccountName -SasToken $effectiveSas -ErrorAction Stop
    Write-Log -Message "Contesto Azure Storage creato tramite SAS." -Level "SUCCESS"

    if ([string]::IsNullOrWhiteSpace($BlobPrefix)) {
        $blobs = @(Get-AzStorageBlob -Container $ContainerName -Context $context -ErrorAction Stop)
    }
    else {
        $blobs = @(Get-AzStorageBlob -Container $ContainerName -Prefix $BlobPrefix -Context $context -ErrorAction Stop)
    }

    Write-Log -Message "Blob individuati: $($blobs.Count)."

    foreach ($blob in $blobs) {
        $name = $blob.Name
        $tmp = $null
        $stage = $null
        $final = $null
        $wasDeleted = $false

        try {
            if ($name.EndsWith('/')) {
                $skipped++
                Write-Log -Message "Directory virtuale ignorata: $name"
                continue
            }

            $stage = Get-SafePath -BasePath $StagingPath -BlobName $name
            $final = Get-SafePath -BasePath $DestinationPath -BlobName $name
            $tmp = "$stage.downloading"

            Ensure-Directory -Path (Split-Path -Path $stage -Parent)
            Ensure-Directory -Path (Split-Path -Path $final -Parent)

            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $stage -Force -ErrorAction SilentlyContinue

            Write-Log -Message "Download: $name"
            Get-AzStorageBlobContent -Container $ContainerName -Blob $name -Destination $tmp -Context $context -Force -ErrorAction Stop | Out-Null

            if (-not (Test-Path -LiteralPath $tmp -PathType Leaf)) {
                throw "File temporaneo non creato."
            }

            $localSize = [int64](Get-Item -LiteralPath $tmp).Length
            $remoteSize = [int64]$blob.Length
            if ($localSize -ne $remoteSize) {
                throw "Dimensione non valida. Azure=$remoteSize byte; locale=$localSize byte."
            }

            Move-Item -LiteralPath $tmp -Destination $stage -Force -ErrorAction Stop
            $downloaded++
            $totalBytes += $localSize
            Write-Log -Message "Download verificato: $name ($localSize byte)." -Level "SUCCESS"

            Remove-AzStorageBlob -Container $ContainerName -Blob $name -Context $context -Force -ErrorAction Stop | Out-Null
            $wasDeleted = $true
            $deleted++
            Write-Log -Message "Blob eliminato da Azure: $name" -Level "SUCCESS"

            if (Test-Path -LiteralPath $final -PathType Leaf) {
                $backup = "$final.replaced.$((Get-Date).ToString('yyyyMMddHHmmssfff'))"
                Move-Item -LiteralPath $final -Destination $backup -Force -ErrorAction Stop
                Write-Log -Message "File locale esistente salvato come backup: $backup" -Level "WARN"
            }

            Move-Item -LiteralPath $stage -Destination $final -Force -ErrorAction Stop
            $published++
            Write-Log -Message "File pubblicato: $final" -Level "SUCCESS"
        }
        catch {
            $failed++
            $exitCode = 1
            Write-Log -Message "Errore su '$name': $($_.Exception.Message)" -Level "ERROR"

            if ($tmp) { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }

            if ($wasDeleted -and $stage -and (Test-Path -LiteralPath $stage -PathType Leaf)) {
                Write-Log -Message "Blob eliminato ma pubblicazione fallita. File integro conservato in staging: $stage" -Level "ERROR"
            }
            elseif (-not $wasDeleted) {
                Write-Log -Message "Il blob resta in Azure e sara ritentato alla prossima esecuzione." -Level "WARN"
            }
        }
    }
}
catch {
    if ($exitCode -eq 0) { $exitCode = 1 }
    Write-Log -Message "Errore generale: $($_.Exception.Message)" -Level "ERROR"
}
finally {
    $duration = (Get-Date) - $startTime
    $mb = [Math]::Round($totalBytes / 1MB, 2)
    Write-Log -Message "File scaricati e verificati: $downloaded"
    Write-Log -Message "Blob eliminati da Azure: $deleted"
    Write-Log -Message "File pubblicati localmente: $published"
    Write-Log -Message "Elementi ignorati: $skipped"
    Write-Log -Message "Elementi in errore: $failed"
    Write-Log -Message "Totale scaricato: $mb MB"
    Write-Log -Message ("Durata: {0:hh\:mm\:ss}" -f $duration)
    Write-Log -Message "Exit code: $exitCode"
    Write-Log -Message "============================================================"

    if ($hasMutex -and $mutex) { $mutex.ReleaseMutex() }
    if ($mutex) { $mutex.Dispose() }
}

exit $exitCode
