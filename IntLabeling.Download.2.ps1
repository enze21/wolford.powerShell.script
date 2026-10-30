#Requires -Version 7.2
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

# ============================================================
# AZURE STORAGE CONFIGURATION
# Replace ONLY $SasToken with the SAS generated for container stealthftp.
# Required SAS permissions: Read + List + Delete (rld), HTTPS only.
# The leading '?' is optional.
# ============================================================
$StorageAccountName = "stowofosftpdev"
$ContainerName      = "stealthftp"
$SasToken           = "sp=racwdl&st=2026-10-29T11:27:35Z&se=2027-11-01T19:42:35Z&spr=https&sv=2026-02-06&sr=c&sig=PugoygceqYO506kEvOnyIaxkhboVZ2NsMLZxMVGAAl4%3D"

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
    param([Parameter(Mandatory)][string]$Path)
    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Write-Log {
    param(
        [Parameter(Mandatory)][string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")][string]$Level = "INFO"
    )
    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"), $Level, $Message
    Write-Host $line
    if ($script:LogFile) {
        Add-Content -LiteralPath $script:LogFile -Value $line -Encoding utf8
    }
}

function Get-SafePath {
    param(
        [Parameter(Mandatory)][string]$BasePath,
        [Parameter(Mandatory)][string]$BlobName
    )

    $base = [IO.Path]::GetFullPath($BasePath)
    $relative = $BlobName.Replace('/', [IO.Path]::DirectorySeparatorChar)
    $result = [IO.Path]::GetFullPath((Join-Path $base $relative))
    $prefix = $base.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar

    if (-not $result.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Percorso locale non valido per il blob '$BlobName'."
    }
    return $result
}

try {
    Ensure-Directory $DestinationPath
    Ensure-Directory $StagingPath
    Ensure-Directory $LogDirectory

    $script:LogFile = Join-Path $LogDirectory ("Download-StealthFTP_{0}.log" -f (Get-Date -Format "yyyyMMdd"))

    $mutex = [Threading.Mutex]::new($false, "Global\Download-StealthFTP_$StorageAccountName`_$ContainerName")
    $hasMutex = $mutex.WaitOne([TimeSpan]::FromSeconds(5))
    if (-not $hasMutex) {
        $exitCode = 2
        throw "Un'altra istanza dello script e' gia in esecuzione."
    }

    Write-Log "============================================================"
    Write-Log "Avvio processo. Storage=$StorageAccountName; Container=$ContainerName"
    Write-Log "Staging=$StagingPath"
    Write-Log "Destination=$DestinationPath"

    if ([string]::IsNullOrWhiteSpace($SasToken) -or $SasToken -eq "PASTE_SAS_TOKEN_HERE") {
        throw "SAS non configurata. Inserire il token nella variabile `$SasToken."
    }

    $effectiveSas = $SasToken.Trim().TrimStart('?')
    $context = New-AzStorageContext -StorageAccountName $StorageAccountName -SasToken $effectiveSas
    Write-Log "Contesto Azure Storage creato tramite SAS." "SUCCESS"

    if ([string]::IsNullOrWhiteSpace($BlobPrefix)) {
        $blobs = @(Get-AzStorageBlob -Container $ContainerName -Context $context)
    }
    else {
        $blobs = @(Get-AzStorageBlob -Container $ContainerName -Prefix $BlobPrefix -Context $context)
    }

    Write-Log "Blob individuati: $($blobs.Count)."

    foreach ($blob in $blobs) {
        $name = $blob.Name
        $tmp = $null
        $stage = $null
        $final = $null
        $wasDeleted = $false

        try {
            if ($name.EndsWith('/')) {
                $skipped++
                Write-Log "Directory virtuale ignorata: $name"
                continue
            }

            $stage = Get-SafePath $StagingPath $name
            $final = Get-SafePath $DestinationPath $name
            $tmp = "$stage.downloading"

            Ensure-Directory (Split-Path $stage -Parent)
            Ensure-Directory (Split-Path $final -Parent)

            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $stage -Force -ErrorAction SilentlyContinue

            Write-Log "Download: $name"
            Get-AzStorageBlobContent -Container $ContainerName -Blob $name -Destination $tmp -Context $context -Force | Out-Null

            if (-not (Test-Path -LiteralPath $tmp -PathType Leaf)) {
                throw "File temporaneo non creato."
            }

            $localSize = [int64](Get-Item -LiteralPath $tmp).Length
            $remoteSize = [int64]$blob.Length

            if ($localSize -ne $remoteSize) {
                throw "Dimensione non valida. Azure=$remoteSize byte; locale=$localSize byte."
            }

            Move-Item -LiteralPath $tmp -Destination $stage -Force
            $downloaded++
            $totalBytes += $localSize
            Write-Log "Download verificato: $name ($localSize byte)." "SUCCESS"

            Remove-AzStorageBlob -Container $ContainerName -Blob $name -Context $context -Force | Out-Null
            $wasDeleted = $true
            $deleted++
            Write-Log "Blob eliminato da Azure: $name" "SUCCESS"

            if (Test-Path -LiteralPath $final -PathType Leaf) {
                $backup = "$final.replaced.$((Get-Date).ToString('yyyyMMddHHmmssfff'))"
                Move-Item -LiteralPath $final -Destination $backup -Force
                Write-Log "File locale esistente salvato come backup: $backup" "WARN"
            }

            Move-Item -LiteralPath $stage -Destination $final -Force
            $published++
            Write-Log "File pubblicato: $final" "SUCCESS"
        }
        catch {
            $failed++
            $exitCode = 1
            Write-Log "Errore su '$name': $($_.Exception.Message)" "ERROR"

            if ($tmp) {
                Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
            }

            if ($wasDeleted -and $stage -and (Test-Path -LiteralPath $stage -PathType Leaf)) {
                Write-Log "Blob eliminato ma pubblicazione fallita. File integro conservato in staging: $stage" "ERROR"
            }
            elseif (-not $wasDeleted) {
                Write-Log "Il blob resta in Azure e sara ritentato alla prossima esecuzione." "WARN"
            }
        }
    }
}
catch {
    if ($exitCode -eq 0) { $exitCode = 1 }
    Write-Log "Errore generale: $($_.Exception.Message)" "ERROR"
}
finally {
    $duration = (Get-Date) - $startTime
    $mb = [Math]::Round($totalBytes / 1MB, 2)

    Write-Log "File scaricati e verificati: $downloaded"
    Write-Log "Blob eliminati da Azure: $deleted"
    Write-Log "File pubblicati localmente: $published"
    Write-Log "Elementi ignorati: $skipped"
    Write-Log "Elementi in errore: $failed"
    Write-Log "Totale scaricato: $mb MB"
    Write-Log ("Durata: {0:hh\:mm\:ss}" -f $duration)
    Write-Log "Exit code: $exitCode"
    Write-Log "============================================================"

    if ($hasMutex -and $mutex) { $mutex.ReleaseMutex() }
    if ($mutex) { $mutex.Dispose() }
}

exit $exitCode
