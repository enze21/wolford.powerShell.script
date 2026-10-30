#Requires -Version 5.1

[CmdletBinding()]
param(
    [string]$RemoteFolder    = "Test/In/IntLabeling",
    [string]$DestinationPath = "D:\StealthFTP\Inbound",
    [string]$StagingPath     = "D:\StealthFTP\Staging",
    [string]$LogDirectory    = "D:\StealthFTP\Logs"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$ProgressPreference = "SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Azure Storage configuration
$StorageAccountName = "stowofosftpdev"
$ContainerName      = "stealthftp"
$SasToken           = "sp=racwdl&st=2026-10-29T11:27:35Z&se=2027-11-01T19:42:35Z&spr=https&sv=2026-02-06&sr=c&sig=PugoygceqYO506kEvOnyIaxkhboVZ2NsMLZxMVGAAl4%3D"

# When true, only files directly inside RemoteFolder are processed.
# Any nested subfolder is ignored.
$OnlyDirectFiles = $true

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
    return (($BlobName -split '/', -1 | ForEach-Object { [Uri]::EscapeDataString($_) }) -join '/')
}

function Get-SasQuery {
    param([Parameter(Mandatory=$true)][string]$Value)
    $valueClean = $Value.Trim()
    if ($valueClean -match '^https://') {
        return (New-Object Uri($valueClean)).Query.TrimStart('?')
    }
    return $valueClean.TrimStart('?')
}

function Get-XmlFromResponse {
    param([Parameter(Mandatory=$true)]$Response)
    $bytes = $Response.RawContentStream.ToArray()
    $text = [Text.Encoding]::UTF8.GetString($bytes)
    $text = $text.TrimStart([char]0xFEFF)
    if ($text.StartsWith("ï»¿")) { $text = $text.Substring(3) }
    $document = New-Object System.Xml.XmlDocument
    $document.PreserveWhitespace = $false
    $document.LoadXml($text)
    return $document
}

function Invoke-ListBlobs {
    param(
        [Parameter(Mandatory=$true)][string]$BaseContainerUrl,
        [Parameter(Mandatory=$true)][string]$SasQuery,
        [Parameter(Mandatory=$true)][string]$Prefix
    )

    $items = New-Object System.Collections.ArrayList
    $marker = ""

    do {
        $query = "restype=container&comp=list&maxresults=5000&prefix=$([Uri]::EscapeDataString($Prefix))"
        if (-not [string]::IsNullOrWhiteSpace($marker)) {
            $query += "&marker=$([Uri]::EscapeDataString($marker))"
        }

        $uri = "$BaseContainerUrl`?$query&$SasQuery"
        $response = Invoke-WebRequest -Method Get -Uri $uri -Headers @{"x-ms-version"="2021-12-02"} -UseBasicParsing -ErrorAction Stop
        $xml = Get-XmlFromResponse -Response $response

        if ($xml.EnumerationResults.Blobs.Blob) {
            foreach ($blob in @($xml.EnumerationResults.Blobs.Blob)) {
                [void]$items.Add($blob)
            }
        }
        $marker = [string]$xml.EnumerationResults.NextMarker
    }
    while (-not [string]::IsNullOrWhiteSpace($marker))

    return $items
}

try {
    Ensure-Directory -Path $DestinationPath
    Ensure-Directory -Path $StagingPath
    Ensure-Directory -Path $LogDirectory

    $script:LogFile = Join-Path $LogDirectory ("IntLabeling.Download_{0}.log" -f (Get-Date -Format "yyyyMMdd"))

    $mutexName = "Global\IntLabelingDownload_{0}_{1}" -f $StorageAccountName, $ContainerName
    $mutex = New-Object System.Threading.Mutex($false, $mutexName)
    $hasMutex = $mutex.WaitOne([TimeSpan]::FromSeconds(5))
    if (-not $hasMutex) {
        $exitCode = 2
        throw "Another instance of the script is already running."
    }

    $normalizedRemoteFolder = $RemoteFolder.Trim().Trim('/')
    if ([string]::IsNullOrWhiteSpace($normalizedRemoteFolder)) {
        throw "RemoteFolder cannot be empty."
    }
    $remotePrefix = "$normalizedRemoteFolder/"

    Write-Log -Message "============================================================"
    Write-Log -Message "Starting REST transfer. Storage=$StorageAccountName; Container=$ContainerName"
    Write-Log -Message "Remote folder: $remotePrefix"
    Write-Log -Message "Only direct files: $OnlyDirectFiles"

    if ([string]::IsNullOrWhiteSpace($SasToken) -or $SasToken -eq "PASTE_SAS_TOKEN_HERE") {
        throw "SAS is not configured. Replace PASTE_SAS_TOKEN_HERE in the script."
    }

    $sasQuery = Get-SasQuery -Value $SasToken
    $containerUrl = "https://$StorageAccountName.blob.core.windows.net/$ContainerName"
    $blobs = @(Invoke-ListBlobs -BaseContainerUrl $containerUrl -SasQuery $sasQuery -Prefix $remotePrefix)
    Write-Log -Message "Items returned for prefix '$remotePrefix': $($blobs.Count)."

    foreach ($blob in $blobs) {
        $blobName = [string]$blob.Name
        $temporaryFile = $null
        $stagedFile = $null
        $finalFile = $null
        $blobDeleted = $false

        try {
            $resourceType = [string]$blob.Properties.ResourceType
            if ([string]::IsNullOrWhiteSpace($blobName) -or $resourceType -eq "directory" -or $blobName.EndsWith('/')) {
                $skipped++
                Write-Log -Message "Directory ignored: $blobName"
                continue
            }

            if (-not $blobName.StartsWith($remotePrefix, [StringComparison]::OrdinalIgnoreCase)) {
                $skipped++
                Write-Log -Message "Blob outside configured folder ignored: $blobName" -Level "WARN"
                continue
            }

            $relativeName = $blobName.Substring($remotePrefix.Length)
            if ([string]::IsNullOrWhiteSpace($relativeName)) {
                $skipped++
                continue
            }

            if ($OnlyDirectFiles -and $relativeName.Contains('/')) {
                $skipped++
                Write-Log -Message "Nested file ignored: $blobName"
                continue
            }

            # Flatten direct files into the local staging and destination folders.
            # If recursive mode is enabled, relative subfolders are preserved.
            $localRelativeName = $relativeName.Replace('/', [IO.Path]::DirectorySeparatorChar)
            $stagedFile = Join-Path $StagingPath $localRelativeName
            $finalFile = Join-Path $DestinationPath $localRelativeName
            $temporaryFile = "$stagedFile.downloading"

            Ensure-Directory -Path (Split-Path $stagedFile -Parent)
            Ensure-Directory -Path (Split-Path $finalFile -Parent)
            Remove-Item -LiteralPath $temporaryFile -Force -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $stagedFile -Force -ErrorAction SilentlyContinue

            $remoteLength = [int64]$blob.Properties.'Content-Length'
            $etag = [string]$blob.Properties.Etag
            $blobUrl = "$containerUrl/$(Encode-BlobPath -BlobName $blobName)`?$sasQuery"

            Write-Log -Message "Downloading: $blobName"
            Invoke-WebRequest -Method Get -Uri $blobUrl -Headers @{"x-ms-version"="2021-12-02"} -OutFile $temporaryFile -UseBasicParsing -ErrorAction Stop

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
            Write-Log -Message "Download verified: $blobName ($localLength bytes)." -Level "SUCCESS"

            $deleteHeaders = @{"x-ms-version"="2021-12-02"}
            if (-not [string]::IsNullOrWhiteSpace($etag)) { $deleteHeaders["If-Match"] = $etag }
            Invoke-RestMethod -Method Delete -Uri $blobUrl -Headers $deleteHeaders -UseBasicParsing -ErrorAction Stop | Out-Null
            $blobDeleted = $true
            $deleted++
            Write-Log -Message "Blob deleted from Azure: $blobName" -Level "SUCCESS"

            if (Test-Path -LiteralPath $finalFile -PathType Leaf) {
                $backup = "$finalFile.replaced.$((Get-Date).ToString('yyyyMMddHHmmssfff'))"
                Move-Item -LiteralPath $finalFile -Destination $backup -Force
                Write-Log -Message "Existing local file moved to backup: $backup" -Level "WARN"
            }

            Move-Item -LiteralPath $stagedFile -Destination $finalFile -Force
            $published++
            Write-Log -Message "File published: $finalFile" -Level "SUCCESS"
        }
        catch {
            $failed++
            $exitCode = 1
            Write-Log -Message "Error processing '$blobName': $($_.Exception.Message)" -Level "ERROR"
            if ($temporaryFile) { Remove-Item -LiteralPath $temporaryFile -Force -ErrorAction SilentlyContinue }

            if ($blobDeleted -and $stagedFile -and (Test-Path -LiteralPath $stagedFile -PathType Leaf)) {
                Write-Log -Message "Blob was deleted but publishing failed. Verified file remains in staging: $stagedFile" -Level "ERROR"
            }
            elseif (-not $blobDeleted) {
                Write-Log -Message "Blob remains in Azure and will be retried." -Level "WARN"
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
    Write-Log -Message "Summary: downloaded=$downloaded; deleted=$deleted; published=$published; skipped=$skipped; errors=$failed; MB=$([Math]::Round($totalBytes/1MB,2)); duration=$($duration.ToString('hh\:mm\:ss')); exitCode=$exitCode"
    Write-Log -Message "============================================================"
    if ($hasMutex -and $mutex) { $mutex.ReleaseMutex() }
    if ($mutex) { $mutex.Dispose() }
}

exit $exitCode
