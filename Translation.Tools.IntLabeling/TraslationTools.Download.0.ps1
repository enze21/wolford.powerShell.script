#Requires -Version 5.0
[CmdletBinding()]
param(
[string]$RemoteFolder="Test/In/TranslationTools",
[string]$DestinationPath="D:\StealthFTP\Inbound\TranslationTools",
[string]$StagingPath="D:\StealthFTP\Staging",
[string]$LogDirectory="D:\StealthFTP\Logs"
)
Set-StrictMode -Version Latest
$ErrorActionPreference="Stop"
$ProgressPreference="SilentlyContinue"
[Net.ServicePointManager]::SecurityProtocol=[Net.SecurityProtocolType]::Tls12
$StorageAccountName="stowofosftpdev"
$ContainerName="stealthftp"
$SasToken="sp=racwdl&st=2026-10-29T11:27:35Z&se=2027-11-01T19:42:35Z&spr=https&sv=2026-02-06&sr=c&sig=PugoygceqYO506kEvOnyIaxkhboVZ2NsMLZxMVGAAl4%3D"
$OnlyDirectFiles=$true
$script:LogFile=$null;
$exitCode=0;
$downloaded=0;
$deleted=0;
$published=0;
$skipped=0;
$failed=0;
$totalBytes=[int64]0;
$startTime=Get-Date;
$mutex=$null;
$hasMutex=$false
function Ensure-Directory([string]$Path){if(-not(Test-Path -LiteralPath $Path -PathType Container)){New-Item -ItemType Directory -Path $Path -Force|Out-Null}}
function Write-Log([string]$Message,[ValidateSet("INFO","WARN","ERROR","SUCCESS")][string]$Level="INFO"){$line="{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss.fff"),$Level,$Message;
    Write-Host $line;
    if($script:LogFile){Add-Content -LiteralPath $script:LogFile -Value $line -Encoding UTF8}}
function Encode-BlobPath([string]$BlobName){return(($BlobName -split '/',-1|ForEach-Object{[Uri]::EscapeDataString($_)})-join '/')}
function Get-SasQuery([string]$Value){$v=$Value.Trim();
    if($v -match '^https://'){return(New-Object Uri($v)).Query.TrimStart('?')};
    return $v.TrimStart('?')}
function Get-XmlFromResponse($Response){$bytes=$Response.RawContentStream.ToArray();
    $text=[Text.Encoding]::UTF8.GetString($bytes);
    $text=$text.TrimStart([char]0xFEFF);
    if($text.StartsWith("ï»¿")){$text=$text.Substring(3)};
    $doc=New-Object System.Xml.XmlDocument;
    $doc.LoadXml($text);
    return $doc}
function ConvertTo-WindowsSafeFileName([string]$Name){
    # Azure Blob names may contain characters invalid in Windows filenames, e.g. ':' in timestamps.
    # Keep the Azure blob name unchanged remotely;
    # sanitize only the local filename.
    $invalid=[IO.Path]::GetInvalidFileNameChars()
    $chars=$Name.ToCharArray()
    for($i=0;
    $i -lt $chars.Length;
    $i++){if($invalid -contains $chars[$i]){$chars[$i]='-'}}
    $safe=(-join $chars).TrimEnd(' ','.')
    if([string]::IsNullOrWhiteSpace($safe)){throw "Blob filename '$Name' cannot be converted to a valid Windows filename."}
    return $safe
}
function Invoke-ListBlobs([string]$BaseContainerUrl,[string]$SasQuery,[string]$Prefix){$items=New-Object System.Collections.ArrayList;
    $marker="";
    do{$q="restype=container&comp=list&maxresults=5000&prefix=$([Uri]::EscapeDataString($Prefix))";
        if($marker){$q+="&marker=$([Uri]::EscapeDataString($marker))"};
        $resp=Invoke-WebRequest -Method Get -Uri "$BaseContainerUrl`?$q&$SasQuery" -Headers @{"x-ms-version"="2021-12-02"} -UseBasicParsing;
        $xml=Get-XmlFromResponse $resp;
        if($xml.EnumerationResults.Blobs.Blob){foreach($b in @($xml.EnumerationResults.Blobs.Blob)){[void]$items.Add($b)}};
        $marker=[string]$xml.EnumerationResults.NextMarker}while($marker);
    return $items}
try{
    Ensure-Directory $DestinationPath;
    Ensure-Directory $StagingPath;
    Ensure-Directory $LogDirectory;
    $script:LogFile=Join-Path $LogDirectory ("TranslationTools.Download_{0}.log" -f(Get-Date -Format yyyyMMdd))
    $mutex=New-Object System.Threading.Mutex($false,"Global\TranslationToolsDownload");
    $hasMutex=$mutex.WaitOne([TimeSpan]::FromSeconds(5));
    if(-not$hasMutex){$exitCode=2;
        throw "Another instance is running."}
    $folder=$RemoteFolder.Trim().Trim('/');
    $prefix="$folder/";
    Write-Log "Starting REST transfer. Storage=$StorageAccountName;
    Container=$ContainerName;
    RemoteFolder=$prefix"
    if([string]::IsNullOrWhiteSpace($SasToken)-or$SasToken -eq "PASTE_SAS_TOKEN_HERE"){throw "SAS is not configured."}
    $sas=Get-SasQuery $SasToken;
    $containerUrl="https://$StorageAccountName.blob.core.windows.net/$ContainerName";
    $blobs=@(Invoke-ListBlobs $containerUrl $sas $prefix);
    Write-Log "Items returned: $($blobs.Count)."
    foreach($blob in $blobs){$blobName=[string]$blob.Name;
        $tmp=$null;
        $stage=$null;
        $final=$null;
        $blobDeleted=$false;
        try{
            $resourceType=[string]$blob.Properties.ResourceType;
            if(-not$blobName-or$resourceType -eq 'directory'-or$blobName.EndsWith('/')){$skipped++;
                continue}
            $relative=$blobName.Substring($prefix.Length);
            if($OnlyDirectFiles -and $relative.Contains('/')){$skipped++;
                Write-Log "Nested file ignored: $blobName";
                continue}
            $safeName=ConvertTo-WindowsSafeFileName $relative
            if($safeName -ne $relative){Write-Log "Windows filename normalized: '$relative' -> '$safeName'" "WARN"}
            $stage=Join-Path $StagingPath $safeName;
            $final=Join-Path $DestinationPath $safeName;
            $tmp="$stage.downloading";
            Remove-Item $tmp -Force -ErrorAction SilentlyContinue;
            Remove-Item $stage -Force -ErrorAction SilentlyContinue
            $remoteLength=[int64]$blob.Properties.'Content-Length';
            $etag=[string]$blob.Properties.Etag;
            $blobUrl="$containerUrl/$(Encode-BlobPath $blobName)`?$sas"
            Write-Log "Downloading: $blobName -> $safeName";
            Invoke-WebRequest -Method Get -Uri $blobUrl -Headers @{"x-ms-version"="2021-12-02"} -OutFile $tmp -UseBasicParsing -ErrorAction Stop
            $localLength=[int64](Get-Item -LiteralPath $tmp).Length;
            if($localLength -ne $remoteLength){throw "Size mismatch. Azure=$remoteLength local=$localLength"};
            Move-Item $tmp $stage -Force;
            $downloaded++;
            $totalBytes+=$localLength;
            Write-Log "Download verified: $safeName" "SUCCESS"
            $headers=@{"x-ms-version"="2021-12-02"};
            if($etag){$headers["If-Match"]=$etag};
            Invoke-RestMethod -Method Delete -Uri $blobUrl -Headers $headers -UseBasicParsing -ErrorAction Stop|Out-Null;
            $blobDeleted=$true;
            $deleted++;
            Write-Log "Blob deleted: $blobName" "SUCCESS"
            if(Test-Path -LiteralPath $final){$backup="$final.replaced.$((Get-Date).ToString('yyyyMMddHHmmssfff'))";
                Move-Item $final $backup -Force;
                Write-Log "Existing file backed up: $backup" "WARN"};
            Move-Item $stage $final -Force;
            $published++;
            Write-Log "File published: $final" "SUCCESS"
        }catch{$failed++;
            $exitCode=1;
            Write-Log "Error processing '$blobName': $($_.Exception.Message)" "ERROR";
            if($tmp){Remove-Item $tmp -Force -ErrorAction SilentlyContinue};
            if($blobDeleted -and $stage -and(Test-Path $stage)){Write-Log "Verified file remains in staging: $stage" "ERROR"}elseif(-not$blobDeleted){Write-Log "Blob remains in Azure for retry." "WARN"}}}
}catch{if($exitCode -eq 0){$exitCode=1};
    Write-Log "General error: $($_.Exception.Message)" "ERROR"}finally{$duration=(Get-Date)-$startTime;
    Write-Log "Summary: downloaded=$downloaded;
    deleted=$deleted;
    published=$published;
    skipped=$skipped;
    errors=$failed;
    MB=$([Math]::Round($totalBytes/1MB,2));
    duration=$($duration.ToString('hh\:mm\:ss'));
    exitCode=$exitCode";
    if($hasMutex -and $mutex){$mutex.ReleaseMutex()};
    if($mutex){$mutex.Dispose()}}
exit $exitCode
