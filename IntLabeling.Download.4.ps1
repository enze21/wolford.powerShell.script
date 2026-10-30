#Requires -Version 5.1



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
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# ============================================================
# AZURE STORAGE CONFIGURATION
# SAS permissions required: Read + List + Delete (rld)
# Paste either the SAS token or the full container SAS URL.
# ============================================================
$StorageAccountName = "stowofosftpdev"
$ContainerName      = "stealthftp"
$SasToken           = "sp=racwdl&st=2026-10-29T11:27:35Z&se=2027-11-01T19:42:35Z&spr=https&sv=2026-02-06&sr=c&sig=PugoygceqYO506kEvOnyIaxkhboVZ2NsMLZxMVGAAl4%3D"

$script:LogFile = $null
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

function Encode-BlobPath {
    param([Parameter(Mandatory=$true)][string]$BlobName)
    $segments = $BlobName -split '/', -1
    $encoded = foreach ($segment in $segments) {
        [Uri]::EscapeDataString($segment)
    }
    return ($encoded -join '/')
}

function Get-SafeLocalPath {
    param(
        [Parameter(Mandatory=$true)][string]$BasePath,
        [Parameter(Mandatory=$true)][string]$BlobName
    )
    $base = [IO.Path]::GetFullPath($BasePath)
    $relative = $BlobName.Replace('/', [IO.Path]::DirectorySeparatorChar)
    $result = [IO.Path]::GetFullPath((Join-Path -Path $base -ChildPath $relative))
    $prefix = $base.TrimEnd([IO.Path]::DirectorySeparatorChar) + [IO.Path]::DirectorySeparatorChar
    if (-not $result.StartsWith($prefix, [StringComparison]::OrdinalIgnoreCase)) {
        throw "Unsafe local path for blob '$BlobName'."
    }
    return $result
}

function Get-SasQuery {
    param([Parameter(Mandatory=$true)][string]$Value)
    $trimmed = $Value.Trim()
    if ($trimmed -match '^https://') {
        $uri = New-Object Uri($trimmed)
        return $uri.Query.TrimStart('?')
    }
    return $trimmed.TrimStart('?')
}

function Invoke-ListBlobs {
    param(
        [Parameter(Mandatory=$true)][string]$BaseContainerUrl,
        [Parameter(Mandatory=$true)][string]$SasQuery,
        [string]$Prefix
    )

    $allBlobs = New-Object System.Collections.ArrayList
    $marker = ""

    do {
        $query = "restype=container&comp=list&maxresults=5000"
        if (-not [string]::IsNullOrWhiteSpace($Prefix)) {
            $query += "&prefix=$([Uri]::EscapeDataString($Prefix))"
        }
        if (-not [string]::IsNullOrWhiteSpace($marker)) {
            $query += "&marker=$([Uri]::EscapeDataString($marker))"
        }

        $uri = "$BaseContainerUrl`?$query&$SasQuery"
        [xml]$xml = Invoke-RestMethod -Method Get -Uri $uri -Headers @{ "x-ms-version" = "2021-12-02" } -UseBasicParsing

        if ($xml.EnumerationResults.Blobs.Blob) {
            foreach ($item in @($xml.EnumerationResults.Blobs.Blob)) {
                [void]$allBlobs.Add($item)
            }
        }
        $marker = [string]$xml.EnumerationResults.NextMarker
    }
    while (-not [string]::IsNullOrWhiteSpace($marker))

    return $allBlobs
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
        throw "Another instance of the script is already running."
    }

    Write-Log -Message "============================================================"
    Write-Log -Message "Starting REST transfer. Storage=$StorageAccountName; Container=$ContainerName"

    if ([string]::IsNullOrWhiteSpace($SasToken) -or $SasToken -eq "PASTE_SAS_TOKEN_HERE") {
        throw "SAS is not configured. Replace PASTE_SAS_TOKEN_HERE in the script."
    }

    $sasQuery = Get-SasQuery -Value $SasToken
    $containerUrl = "https://$StorageAccountName.blob.core.windows.net/$ContainerName"
    $blobs = @(Invoke-ListBlobs -BaseContainerUrl $containerUrl -SasQuery $sasQuery -Prefix $BlobPrefix)
    Write-Log -Message "Blobs found: $($blobs.Count)."

    foreach ($blob in $blobs) {
        $name = [string]$blob.Name
        $temporaryFile = $null
        $stagedFile = $null
        $finalFile = $null
        $blobDeleted = $false

        try {
            if ([string]::IsNullOrWhiteSpace($name) -or $name.EndsWith('/')) {
                $skipped++
                Write-Log -Message "Virtual directory ignored: $name"
                continue
            }

            $remoteLength = [int64]$blob.Properties.'Content-Length'
            $etag = [string]$blob.Properties.Etag
            $encodedBlobName = Encode-BlobPath -BlobName $name
            $blobUrl = "$containerUrl/$encodedBlobName`?$sasQuery"

            $stagedFile = Get-SafeLocalPath -BasePath $StagingPath -BlobName $name
            $finalFile = Get-SafeLocalPath -BasePath $DestinationPath -BlobName $name
            $temporaryFile = "$stagedFile.downloading"

            Ensure-Directory -Path (Split-Path -Path $stagedFile -Parent)
            Ensure-Directory -Path (Split-Path -Path $finalFile -Parent)
            Remove-Item -LiteralPath $temporaryFile -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $stagedFile -Force -ErrorAction SilentlyContinue

            Write-Log -Message "Downloading: $name"
            Invoke-WebRequest -Method Get -Uri $blobUrl -Headers @{ "x-ms-version" = "2021-12-02" } -OutFile $temporaryFile -UseBasicParsing

            if (-not (Test-Path -LiteralPath $temporaryFile -PathType Leaf)) {
                throw "Temporary file was not created."
            }

            $localLength = [int64](Get-Item -LiteralPath $temporaryFile).Length
            if ($localLength -ne $remoteLength) {
                throw "Size verification failed. Azure=$remoteLength bytes; local=$localLength bytes."
            }

            Move-Item -LiteralPath $temporaryFile -Destination $stagedFile -Force
            $downloaded++
            $totalBytes += $localLength
            Write-Log -Message "Download verified: $name ($localLength bytes)." -Level "SUCCESS"

            $deleteHeaders = @{ "x-ms-version" = "2021-12-02" }
            if (-not [string]::IsNullOrWhiteSpace($etag)) {
                $deleteHeaders["If-Match"] = $etag
            }
            Invoke-RestMethod -Method Delete -Uri $blobUrl -Headers $deleteHeaders -UseBasicParsing | Out-Null
            $blobDeleted = $true
            $deleted++
            Write-Log -Message "Blob deleted from Azure: $name" -Level "SUCCESS"

            if (Test-Path -LiteralPath $finalFile -PathType Leaf) {
                $backup = "$finalFile.replaced.$((Get-Date).ToString('yyyyMMddHHmmssfff'))"
                Move-Item -LiteralPath $finalFile -Destination $backup -Force
                Write-Log -Message "Existing local file moved to: $backup" -Level "WARN"
            }

            Move-Item -LiteralPath $stagedFile -Destination $finalFile -Force
            $published++
            Write-Log -Message "File published: $finalFile" -Level "SUCCESS"
        }
        catch {
            $failed++
            $exitCode = 1
            Write-Log -Message "Error processing '$name': $($_.Exception.Message)" -Level "ERROR"

            if ($temporaryFile) {
                Remove-Item -LiteralPath $temporaryFile -Force -ErrorAction SilentlyContinue
            }

            if ($blobDeleted -and $stagedFile -and (Test-Path -LiteralPath $stagedFile -PathType Leaf)) {
                Write-Log -Message "Blob was deleted, but publishing failed. Verified file remains in staging: $stagedFile" -Level "ERROR"
            }
            elseif (-not $blobDeleted) {
                Write-Log -Message "Blob remains in Azure and will be retried on the next run." -Level "WARN"
            }
        }
    }
}
catch {
    if ($exitCode -eq 0) { $exitCode = 1 }
    Write-Log -Message "General error: $($_.Exception.Message)" -Level "ERROR"
}
finally {
    $duration = (Get-Date) - $startTime
    $megabytes = [Math]::Round($totalBytes / 1MB, 2)
    Write-Log -Message "Downloaded and verified: $downloaded"
    Write-Log -Message "Deleted from Azure: $deleted"
    Write-Log -Message "Published locally: $published"
    Write-Log -Message "Skipped: $skipped"
    Write-Log -Message "Errors: $failed"
    Write-Log -Message "Downloaded: $megabytes MB"
    Write-Log -Message ("Duration: {0:hh\:mm\:ss}" -f $duration)
    Write-Log -Message "Exit code: $exitCode"
    Write-Log -Message "============================================================"

    if ($hasMutex -and $mutex) { $mutex.ReleaseMutex() }
    if ($mutex) { $mutex.Dispose() }
}

exit $exitCode
