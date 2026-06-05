#requires -Modules Az.Accounts
<#
Script unico che:
- controlla Queue e Topic/Subscription per un elenco di entity names
- se la DLQ contiene messaggi, esegue l'export dei messaggi
- esporta su CSV e opzionalmente su SQL Server on-prem

NOTE:
- Mantiene la logica dei tre script originali in un unico file.
- Per i Topic controlla tutte le subscription esistenti del topic.
- Per le Queue controlla la queue con lo stesso nome dell'entity.
- Legge i messaggi dalla DLQ in PeekLock e li sblocca a fine elaborazione.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ============================================================
# CONFIG
# ============================================================
$configPath = 'D:\PowerShell\Azure.Queue.Topic\config.json'
if (!(Test-Path $configPath)) {
    Write-Error "Config file not found: $configPath"
    exit 1
}

$config = Get-Content $configPath -Raw | ConvertFrom-Json

$TenantId       = $config.TEST.TenantId
$SubscriptionId = $config.TEST.SubscriptionId
$ResourceGroup  = $config.TEST.ResourceGroup
$Namespace      = $config.TEST.Namespace
$BasePath       = $config.TEST.BasePath

$SqlServer                = $config.TEST.SqlServer
$SqlDatabase              = if ($config.TEST.SqlDatabase) { $config.TEST.SqlDatabase } else { 'WIL.Queue.Topic' }
$SqlSchema                = if ($config.TEST.SqlSchema) { $config.TEST.SqlSchema } else { 'dbo' }
$SqlUseIntegratedSecurity = if ($null -ne $config.TEST.SqlUseIntegratedSecurity) { [bool]$config.TEST.SqlUseIntegratedSecurity } else { $true }
$SqlUser                  = $config.TEST.SqlUser
$SqlPassword              = $config.TEST.SqlPassword

$DateStamp = Get-Date -Format 'dd.MM.yyyy'
$OutputCsv = Join-Path $BasePath ("{0}.ServiceBus_DLQ_Report.csv" -f $DateStamp)
$GlobalLog = Join-Path $BasePath ("{0}.ServiceBus_DLQ_Report.log.txt" -f $DateStamp)
$ReceiveTimeoutSeconds = 3
$RenewBufferSeconds    = 10
$EnableCsvExport       = $true
$EnableSqlExport       = $true

$TargetEntities = @(
    'InventoryAdjustmentCreated',
    'InventoryReceivedCreated',
    'D2SInventoryTransferCreated',
    'S2SInventoryTransferCreated',
    'ReturnOrderCreated',
    'SalesOrderCreated'
)

# ============================================================
# FUNZIONI DI UTILITÀ
# ============================================================
function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $false)][string]$Level = 'INFO',
        [Parameter(Mandatory = $false)][string]$LogFile = $GlobalLog
    )

    $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $line = "$timestamp [$Level] - $Message"
    Write-Host $line

    if ($LogFile) {
        $folder = Split-Path -Parent $LogFile
        if ($folder -and !(Test-Path $folder)) {
            New-Item -ItemType Directory -Path $folder -Force | Out-Null
        }
        Add-Content -Path $LogFile -Value $line
    }
}

function Convert-SecureStringToPlainText {
    param([Parameter(Mandatory = $true)][System.Security.SecureString]$SecureString)

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}

function Get-ServiceBusTokenInfo {
    $tok = Get-AzAccessToken -ResourceUrl 'https://servicebus.azure.net/'
    $plain = Convert-SecureStringToPlainText -SecureString $tok.Token

    [pscustomobject]@{
        AccessToken = $plain
        ExpiresOn   = $tok.ExpiresOn
    }
}

function Get-ValidServiceBusToken {
    param($CurrentTokenInfo)

    if ($null -eq $CurrentTokenInfo) {
        return Get-ServiceBusTokenInfo
    }

    $now = Get-Date
    if ($CurrentTokenInfo.ExpiresOn -le $now.AddMinutes(5)) {
        return Get-ServiceBusTokenInfo
    }

    return $CurrentTokenInfo
}

function Get-HeaderValue {
    param(
        [Parameter(Mandatory = $true)]$Response,
        [Parameter(Mandatory = $true)][string[]]$Names
    )

    foreach ($name in $Names) {
        if ($Response.Headers.Contains($name)) {
            return ($Response.Headers.GetValues($name) | Select-Object -First 1)
        }
        if ($Response.Content.Headers.Contains($name)) {
            return ($Response.Content.Headers.GetValues($name) | Select-Object -First 1)
        }
    }

    return $null
}

function Get-CaseInsensitiveProperty {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string[]]$Names
    )

    if ($null -eq $Object) { return $null }

    foreach ($name in $Names) {
        $prop = $Object.PSObject.Properties | Where-Object { $_.Name -ieq $name } | Select-Object -First 1
        if ($prop) {
            return $prop.Value
        }
    }

    return $null
}

function Invoke-ServiceBusRequest {
    param(
        [Parameter(Mandatory = $true)][System.Net.Http.HttpClient]$HttpClient,
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$AccessToken,
        [Parameter(Mandatory = $false)][string]$Body,
        [Parameter(Mandatory = $false)][string]$ContentType = 'application/atom+xml;type=entry;charset=utf-8'
    )

    $request = [System.Net.Http.HttpRequestMessage]::new([System.Net.Http.HttpMethod]::new($Method), $Uri)
    $request.Headers.TryAddWithoutValidation('Authorization', "Bearer $AccessToken") | Out-Null
    $request.Headers.TryAddWithoutValidation('Accept', 'application/atom+xml,application/json,text/plain,*/*') | Out-Null

    if ($PSBoundParameters.ContainsKey('Body')) {
        $request.Content = [System.Net.Http.StringContent]::new($Body, [System.Text.Encoding]::UTF8, $ContentType)
    }

    return $HttpClient.SendAsync($request).GetAwaiter().GetResult()
}

function Ensure-AzureLogin {
    Write-Log "Login su Azure Tenant $TenantId"
    try {
        $ctx = Get-AzContext -ErrorAction SilentlyContinue
        if (-not $ctx) {
            Connect-AzAccount -Tenant $TenantId -Subscription $SubscriptionId | Out-Null
        }
        else {
            Write-Log "Sessione Azure già attiva: $($ctx.Account.Id)"
            if ($ctx.Subscription -and $ctx.Subscription.Id -ne $SubscriptionId) {
                Set-AzContext -SubscriptionId $SubscriptionId | Out-Null
            }
        }
    }
    catch {
        Write-Log "Errore durante login/context: $($_.Exception.Message)" 'ERROR'
        throw
    }
}

# ============================================================
# FUNZIONI SQL
# ============================================================
function Quote-SqlIdentifier {
    param([Parameter(Mandatory = $true)][string]$Name)
    return '[' + ($Name -replace ']', ']]') + ']'
}

function Get-SqlConnectionString {
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$Database,
        [Parameter(Mandatory = $true)][bool]$UseIntegratedSecurity,
        [Parameter(Mandatory = $false)][string]$User,
        [Parameter(Mandatory = $false)][string]$Password
    )

    if ($UseIntegratedSecurity) {
        return "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;"
    }

    return "Server=$Server;Database=$Database;User ID=$User;Password=$Password;TrustServerCertificate=True;"
}

function Convert-ToNullableDateTime {
    param([Parameter(Mandatory = $false)]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return [DBNull]::Value }
    try { return [datetime]$Value } catch { return [DBNull]::Value }
}

function Convert-ToNullableInt64 {
    param([Parameter(Mandatory = $false)]$Value)
    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) { return [DBNull]::Value }
    try { return [int64]$Value } catch { return [DBNull]::Value }
}

function Ensure-SqlDatabaseAndTable {
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$DatabaseName,
        [Parameter(Mandatory = $true)][string]$SchemaName,
        [Parameter(Mandatory = $true)][string]$TableName,
        [Parameter(Mandatory = $true)][ValidateSet('Queue','TopicSubscription')][string]$Mode,
        [Parameter(Mandatory = $true)][bool]$UseIntegratedSecurity,
        [Parameter(Mandatory = $false)][string]$User,
        [Parameter(Mandatory = $false)][string]$Password
    )

    Add-Type -AssemblyName System.Data

    $masterConn = [System.Data.SqlClient.SqlConnection]::new((Get-SqlConnectionString -Server $Server -Database 'master' -UseIntegratedSecurity $UseIntegratedSecurity -User $User -Password $Password))
    try {
        $masterConn.Open()
        $cmd = $masterConn.CreateCommand()
        $cmd.CommandText = @"
IF DB_ID(@dbName) IS NULL
BEGIN
    DECLARE @sql nvarchar(max);
    SET @sql = N'CREATE DATABASE ' + QUOTENAME(@dbName);
    EXEC(@sql);
END
"@
        [void]$cmd.Parameters.Add('@dbName', [System.Data.SqlDbType]::NVarChar, 128)
        $cmd.Parameters['@dbName'].Value = $DatabaseName
        [void]$cmd.ExecuteNonQuery()
    }
    finally {
        $masterConn.Close()
        $masterConn.Dispose()
    }

    $dbConn = [System.Data.SqlClient.SqlConnection]::new((Get-SqlConnectionString -Server $Server -Database $DatabaseName -UseIntegratedSecurity $UseIntegratedSecurity -User $User -Password $Password))
    try {
        $dbConn.Open()
        $cmd = $dbConn.CreateCommand()
        $schemaQuoted = Quote-SqlIdentifier $SchemaName
        $tableQuoted = Quote-SqlIdentifier $TableName
        $fullTable = "$schemaQuoted.$tableQuoted"

        if ($Mode -eq 'Queue') {
            $createTableSql = @"
IF SCHEMA_ID(@schemaName) IS NULL
BEGIN
    EXEC('CREATE SCHEMA ' + QUOTENAME(@schemaName));
END

IF NOT EXISTS (
    SELECT 1
    FROM sys.tables t
    INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE t.name = @tableName
      AND s.name = @schemaName
)
BEGIN
    CREATE TABLE $fullTable (
        [Id] BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [MessageId] NVARCHAR(255) NOT NULL,
        [SequenceNumber] BIGINT NULL,
        [EnqueuedTimeUtc] DATETIME2(7) NULL,
        [DeadLetterReason] NVARCHAR(1024) NULL,
        [DeadLetterDescription] NVARCHAR(MAX) NULL,
        [MessageBody] NVARCHAR(MAX) NULL,
        [ReadAtUtc] DATETIME2(7) NOT NULL,
        [InsertedAtUtc] DATETIME2(7) NOT NULL DEFAULT SYSUTCDATETIME()
    );
END
"@
        }
        else {
            $createTableSql = @"
IF SCHEMA_ID(@schemaName) IS NULL
BEGIN
    EXEC('CREATE SCHEMA ' + QUOTENAME(@schemaName));
END

IF NOT EXISTS (
    SELECT 1
    FROM sys.tables t
    INNER JOIN sys.schemas s ON t.schema_id = s.schema_id
    WHERE t.name = @tableName
      AND s.name = @schemaName
)
BEGIN
    CREATE TABLE $fullTable (
        [Id] BIGINT IDENTITY(1,1) NOT NULL PRIMARY KEY,
        [TopicName] NVARCHAR(255) NOT NULL,
        [SubscriptionName] NVARCHAR(255) NOT NULL,
        [MessageId] NVARCHAR(255) NOT NULL,
        [SequenceNumber] BIGINT NULL,
        [EnqueuedTimeUtc] DATETIME2(7) NULL,
        [DeadLetterReason] NVARCHAR(1024) NULL,
        [DeadLetterDescription] NVARCHAR(MAX) NULL,
        [MessageBody] NVARCHAR(MAX) NULL,
        [ReadAtUtc] DATETIME2(7) NOT NULL,
        [InsertedAtUtc] DATETIME2(7) NOT NULL DEFAULT SYSUTCDATETIME()
    );
END
"@
        }

        $cmd.CommandText = $createTableSql
        [void]$cmd.Parameters.Add('@tableName', [System.Data.SqlDbType]::NVarChar, 128)
        [void]$cmd.Parameters.Add('@schemaName', [System.Data.SqlDbType]::NVarChar, 128)
        $cmd.Parameters['@tableName'].Value = $TableName
        $cmd.Parameters['@schemaName'].Value = $SchemaName
        [void]$cmd.ExecuteNonQuery()
    }
    finally {
        $dbConn.Close()
        $dbConn.Dispose()
    }
}

function Export-RowsToSql {
    param(
        [Parameter(Mandatory = $true)]$Rows,
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$DatabaseName,
        [Parameter(Mandatory = $true)][string]$SchemaName,
        [Parameter(Mandatory = $true)][string]$TableName,
        [Parameter(Mandatory = $true)][ValidateSet('Queue','TopicSubscription')][string]$Mode,
        [Parameter(Mandatory = $true)][bool]$UseIntegratedSecurity,
        [Parameter(Mandatory = $false)][string]$User,
        [Parameter(Mandatory = $false)][string]$Password
    )

    if ($null -eq $Rows -or $Rows.Count -eq 0) { return }

    Add-Type -AssemblyName System.Data
    $conn = [System.Data.SqlClient.SqlConnection]::new((Get-SqlConnectionString -Server $Server -Database $DatabaseName -UseIntegratedSecurity $UseIntegratedSecurity -User $User -Password $Password))

    try {
        $conn.Open()
        $schemaQuoted = Quote-SqlIdentifier $SchemaName
        $tableQuoted = Quote-SqlIdentifier $TableName
        $fullTable = "$schemaQuoted.$tableQuoted"

        foreach ($row in $Rows) {
            $cmd = $conn.CreateCommand()
            if ($Mode -eq 'Queue') {
                $cmd.CommandText = @"
INSERT INTO $fullTable (
    [MessageId],
    [SequenceNumber],
    [EnqueuedTimeUtc],
    [DeadLetterReason],
    [DeadLetterDescription],
    [MessageBody],
    [ReadAtUtc]
)
VALUES (
    @MessageId,
    @SequenceNumber,
    @EnqueuedTimeUtc,
    @DeadLetterReason,
    @DeadLetterDescription,
    @MessageBody,
    @ReadAtUtc
)
"@
            }
            else {
                $cmd.CommandText = @"
INSERT INTO $fullTable (
    [TopicName],
    [SubscriptionName],
    [MessageId],
    [SequenceNumber],
    [EnqueuedTimeUtc],
    [DeadLetterReason],
    [DeadLetterDescription],
    [MessageBody],
    [ReadAtUtc]
)
VALUES (
    @TopicName,
    @SubscriptionName,
    @MessageId,
    @SequenceNumber,
    @EnqueuedTimeUtc,
    @DeadLetterReason,
    @DeadLetterDescription,
    @MessageBody,
    @ReadAtUtc
)
"@
                [void]$cmd.Parameters.Add('@TopicName', [System.Data.SqlDbType]::NVarChar, 255)
                [void]$cmd.Parameters.Add('@SubscriptionName', [System.Data.SqlDbType]::NVarChar, 255)
                $cmd.Parameters['@TopicName'].Value = [string]$row.TopicName
                $cmd.Parameters['@SubscriptionName'].Value = [string]$row.SubscriptionName
            }

            [void]$cmd.Parameters.Add('@MessageId', [System.Data.SqlDbType]::NVarChar, 255)
            [void]$cmd.Parameters.Add('@SequenceNumber', [System.Data.SqlDbType]::BigInt)
            [void]$cmd.Parameters.Add('@EnqueuedTimeUtc', [System.Data.SqlDbType]::DateTime2)
            [void]$cmd.Parameters.Add('@DeadLetterReason', [System.Data.SqlDbType]::NVarChar, 1024)
            [void]$cmd.Parameters.Add('@DeadLetterDescription', [System.Data.SqlDbType]::NVarChar, -1)
            [void]$cmd.Parameters.Add('@MessageBody', [System.Data.SqlDbType]::NVarChar, -1)
            [void]$cmd.Parameters.Add('@ReadAtUtc', [System.Data.SqlDbType]::DateTime2)

            $cmd.Parameters['@MessageId'].Value = if ($row.MessageId) { [string]$row.MessageId } else { '' }
            $cmd.Parameters['@SequenceNumber'].Value = Convert-ToNullableInt64 $row.SequenceNumber
            $cmd.Parameters['@EnqueuedTimeUtc'].Value = Convert-ToNullableDateTime $row.EnqueuedTimeUtc
            $cmd.Parameters['@DeadLetterReason'].Value = if ($row.DeadLetterReason) { [string]$row.DeadLetterReason } else { [DBNull]::Value }
            $cmd.Parameters['@DeadLetterDescription'].Value = if ($row.DeadLetterDescription) { [string]$row.DeadLetterDescription } else { [DBNull]::Value }
            $cmd.Parameters['@MessageBody'].Value = if ($row.MessageBody) { [string]$row.MessageBody } else { [DBNull]::Value }
            $cmd.Parameters['@ReadAtUtc'].Value = [datetime]$row.ReadAtUtc
            [void]$cmd.ExecuteNonQuery()
            $cmd.Dispose()
        }
    }
    finally {
        $conn.Close()
        $conn.Dispose()
    }
}

# ============================================================
# FUNZIONI DLQ EXPORT
# ============================================================
function Get-QueueDlqCount {
    param([Parameter(Mandatory = $true)][string]$QueueName)

    $queue = Get-AzServiceBusQueue -ResourceGroupName $ResourceGroup -NamespaceName $Namespace -Name $QueueName -ErrorAction SilentlyContinue
    if (-not $queue) { return $null }

    $count = Get-CaseInsensitiveProperty -Object $queue -Names @(
        'DeadLetterMessageCount',
        'CountDetails.DeadLetterMessageCount'
    )

    if ($null -eq $count) {
        # fallback esplicito
        $countDetails = Get-CaseInsensitiveProperty -Object $queue -Names @('CountDetails')
        if ($countDetails) {
            $count = Get-CaseInsensitiveProperty -Object $countDetails -Names @('DeadLetterMessageCount')
        }
    }

    return [pscustomobject]@{
        Exists                  = $true
        Queue                   = $queue
        DeadLetterMessageCount  = [int64]($count | ForEach-Object { if ($_ -eq $null) { 0 } else { $_ } })
    }
}

function Get-TopicSubscriptionsDlqCount {
    param([Parameter(Mandatory = $true)][string]$TopicName)

    $topic = Get-AzServiceBusTopic -ResourceGroupName $ResourceGroup -NamespaceName $Namespace -Name $TopicName -ErrorAction SilentlyContinue
    if (-not $topic) { return @() }

    $subs = Get-AzServiceBusSubscription -ResourceGroupName $ResourceGroup -NamespaceName $Namespace -TopicName $TopicName -ErrorAction SilentlyContinue
    $result = @()

    foreach ($sub in @($subs)) {
        $count = Get-CaseInsensitiveProperty -Object $sub -Names @(
            'DeadLetterMessageCount',
            'CountDetails.DeadLetterMessageCount'
        )
        if ($null -eq $count) {
            $countDetails = Get-CaseInsensitiveProperty -Object $sub -Names @('CountDetails')
            if ($countDetails) {
                $count = Get-CaseInsensitiveProperty -Object $countDetails -Names @('DeadLetterMessageCount')
            }
        }

        $result += [pscustomobject]@{
            TopicName              = $TopicName
            SubscriptionName       = (Get-CaseInsensitiveProperty -Object $sub -Names @('Name'))
            DeadLetterMessageCount = [int64]($count | ForEach-Object { if ($_ -eq $null) { 0 } else { $_ } })
            Subscription           = $sub
        }
    }

    return $result
}

function Read-DlqMessages {
    param(
        [Parameter(Mandatory = $true)][string]$EntityPath,
        [Parameter(Mandatory = $true)][System.Net.Http.HttpClient]$HttpClient,
        [Parameter(Mandatory = $false)][hashtable]$StaticFields = @{},
        [Parameter(Mandatory = $false)][string]$EntityLabel = $EntityPath,
        [Parameter(Mandatory = $false)][string]$EntityLogFile = $GlobalLog
    )

    $rows = New-Object System.Collections.Generic.List[object]
    $lockedMessages = New-Object System.Collections.Generic.List[object]
    $tokenInfo = $null

    $dlqEntityPath = "$EntityPath/`$deadletterqueue"
    $baseUri = "https://$Namespace.servicebus.windows.net/$dlqEntityPath"
    $headUri = "$baseUri/messages/head?timeout=$ReceiveTimeoutSeconds"

    Write-Log "DLQ path REST: $dlqEntityPath" 'INFO' $EntityLogFile

    try {
        while ($true) {
            $tokenInfo = Get-ValidServiceBusToken -CurrentTokenInfo $tokenInfo
            $response = Invoke-ServiceBusRequest -HttpClient $HttpClient -Method 'POST' -Uri $headUri -AccessToken $tokenInfo.AccessToken

            if ($response.StatusCode -eq [System.Net.HttpStatusCode]::NotFound) {
                break
            }

            if (($response.StatusCode.value__ -eq 204) -or ($response.StatusCode -eq [System.Net.HttpStatusCode]::NoContent)) {
                break
            }

            if (-not $response.IsSuccessStatusCode) {
                $respBody = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                throw "Errore lettura DLQ $EntityLabel: HTTP $([int]$response.StatusCode) - $respBody"
            }

            $messageBody = $response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
            $brokerPropsRaw = Get-HeaderValue -Response $response -Names @('BrokerProperties')
            $location = Get-HeaderValue -Response $response -Names @('Location')
            $lockToken = Get-HeaderValue -Response $response -Names @('LockToken')

            $brokerProps = $null
            if ($brokerPropsRaw) {
                try {
                    $brokerProps = $brokerPropsRaw | ConvertFrom-Json
                }
                catch {
                    $brokerProps = $null
                }
            }

            $row = [ordered]@{}
            foreach ($key in $StaticFields.Keys) {
                $row[$key] = $StaticFields[$key]
            }

            $row['MessageId']             = if ($brokerProps) { Get-CaseInsensitiveProperty -Object $brokerProps -Names @('MessageId') } else { $null }
            $row['SequenceNumber']        = if ($brokerProps) { Get-CaseInsensitiveProperty -Object $brokerProps -Names @('SequenceNumber') } else { $null }
            $row['EnqueuedTimeUtc']       = if ($brokerProps) { Get-CaseInsensitiveProperty -Object $brokerProps -Names @('EnqueuedTimeUtc') } else { $null }
            $row['DeadLetterReason']      = if ($brokerProps) { Get-CaseInsensitiveProperty -Object $brokerProps -Names @('DeadLetterReason') } else { $null }
            $row['DeadLetterDescription'] = if ($brokerProps) { Get-CaseInsensitiveProperty -Object $brokerProps -Names @('DeadLetterErrorDescription', 'DeadLetterDescription') } else { $null }
            $row['MessageBody']           = $messageBody
            $row['ReadAtUtc']             = (Get-Date).ToUniversalTime()

            $rows.Add([pscustomobject]$row)

            if ($location) {
                $lockedMessages.Add([pscustomobject]@{
                    Location = $location
                    LockToken = $lockToken
                })
            }

            Write-Log ("Messaggio letto da DLQ {0}. Totale finora: {1}" -f $EntityLabel, $rows.Count) 'INFO' $EntityLogFile
        }
    }
    finally {
        if ($lockedMessages.Count -gt 0) {
            Write-Log ("Avvio unlock finale dei messaggi lockati per {0}" -f $EntityLabel) 'INFO' $EntityLogFile
            foreach ($locked in $lockedMessages) {
                try {
                    $tokenInfo = Get-ValidServiceBusToken -CurrentTokenInfo $tokenInfo
                    # Rilascia il lock senza completare il messaggio
                    $unlockResponse = Invoke-ServiceBusRequest -HttpClient $HttpClient -Method 'PUT' -Uri $locked.Location -AccessToken $tokenInfo.AccessToken -Body ''
                    if (-not $unlockResponse.IsSuccessStatusCode) {
                        Write-Log ("Unlock non riuscito per {0}: HTTP {1}" -f $EntityLabel, [int]$unlockResponse.StatusCode) 'WARN' $EntityLogFile
                    }
                }
                catch {
                    Write-Log ("Errore unlock finale per {0}: {1}" -f $EntityLabel, $_.Exception.Message) 'WARN' $EntityLogFile
                }
            }
        }
    }

    return $rows
}

function Export-QueueDlq {
    param(
        [Parameter(Mandatory = $true)][string]$QueueName,
        [Parameter(Mandatory = $true)][int64]$DeadLetterMessageCount,
        [Parameter(Mandatory = $true)][System.Net.Http.HttpClient]$HttpClient
    )

    $entityLog = Join-Path $BasePath ("{0}.{1}.log.txt" -f $DateStamp, $QueueName)
    $entityCsv = Join-Path $BasePath ("{0}.{1}.csv" -f $DateStamp, $QueueName)

    if (Test-Path $entityLog) { Remove-Item $entityLog -Force }

    Write-Log "Avvio export DLQ Queue: $QueueName (messaggi DLQ: $DeadLetterMessageCount)" 'INFO' $entityLog

    $rows = Read-DlqMessages -EntityPath $QueueName -HttpClient $HttpClient -EntityLabel $QueueName -EntityLogFile $entityLog

    if ($EnableCsvExport -and $rows.Count -gt 0) {
        $rows | Export-Csv -Path $entityCsv -NoTypeInformation -Delimiter ';' -Encoding UTF8
        Write-Log "CSV esportato: $entityCsv" 'INFO' $entityLog
    }

    if ($EnableSqlExport -and $rows.Count -gt 0) {
        Ensure-SqlDatabaseAndTable -Server $SqlServer -DatabaseName $SqlDatabase -SchemaName $SqlSchema -TableName $QueueName -Mode 'Queue' -UseIntegratedSecurity $SqlUseIntegratedSecurity -User $SqlUser -Password $SqlPassword
        Export-RowsToSql -Rows $rows -Server $SqlServer -DatabaseName $SqlDatabase -SchemaName $SqlSchema -TableName $QueueName -Mode 'Queue' -UseIntegratedSecurity $SqlUseIntegratedSecurity -User $SqlUser -Password $SqlPassword
        Write-Log "Export SQL completato su tabella: $SqlSchema.$QueueName" 'INFO' $entityLog
    }

    return [pscustomobject]@{
        EntityType             = 'Queue'
        EntityName             = $QueueName
        SubscriptionName       = $null
        DeadLetterMessageCount = $DeadLetterMessageCount
        ExportedRows           = $rows.Count
        CsvPath                = if ($rows.Count -gt 0) { $entityCsv } else { $null }
        LogPath                = $entityLog
        SqlTable               = if ($rows.Count -gt 0) { "$SqlSchema.$QueueName" } else { $null }
    }
}

function Export-TopicSubscriptionDlq {
    param(
        [Parameter(Mandatory = $true)][string]$TopicName,
        [Parameter(Mandatory = $true)][string]$SubscriptionName,
        [Parameter(Mandatory = $true)][int64]$DeadLetterMessageCount,
        [Parameter(Mandatory = $true)][System.Net.Http.HttpClient]$HttpClient
    )

    $tableName = "$TopicName.$SubscriptionName"
    $entityLog = Join-Path $BasePath ("{0}.{1}.{2}.log.txt" -f $DateStamp, $TopicName, $SubscriptionName)
    $entityCsv = Join-Path $BasePath ("{0}.{1}.{2}.csv" -f $DateStamp, $TopicName, $SubscriptionName)

    if (Test-Path $entityLog) { Remove-Item $entityLog -Force }

    Write-Log "Avvio export DLQ Topic/Subscription: $TopicName / $SubscriptionName (messaggi DLQ: $DeadLetterMessageCount)" 'INFO' $entityLog

    $rows = Read-DlqMessages -EntityPath "$TopicName/Subscriptions/$SubscriptionName" -HttpClient $HttpClient -StaticFields @{ TopicName = $TopicName; SubscriptionName = $SubscriptionName } -EntityLabel "$TopicName/$SubscriptionName" -EntityLogFile $entityLog

    if ($EnableCsvExport -and $rows.Count -gt 0) {
        $rows | Export-Csv -Path $entityCsv -NoTypeInformation -Delimiter ';' -Encoding UTF8
        Write-Log "CSV esportato: $entityCsv" 'INFO' $entityLog
    }

    if ($EnableSqlExport -and $rows.Count -gt 0) {
        Ensure-SqlDatabaseAndTable -Server $SqlServer -DatabaseName $SqlDatabase -SchemaName $SqlSchema -TableName $tableName -Mode 'TopicSubscription' -UseIntegratedSecurity $SqlUseIntegratedSecurity -User $SqlUser -Password $SqlPassword
        Export-RowsToSql -Rows $rows -Server $SqlServer -DatabaseName $SqlDatabase -SchemaName $SqlSchema -TableName $tableName -Mode 'TopicSubscription' -UseIntegratedSecurity $SqlUseIntegratedSecurity -User $SqlUser -Password $SqlPassword
        Write-Log "Export SQL completato su tabella: $SqlSchema.$tableName" 'INFO' $entityLog
    }

    return [pscustomobject]@{
        EntityType             = 'TopicSubscription'
        EntityName             = $TopicName
        SubscriptionName       = $SubscriptionName
        DeadLetterMessageCount = $DeadLetterMessageCount
        ExportedRows           = $rows.Count
        CsvPath                = if ($rows.Count -gt 0) { $entityCsv } else { $null }
        LogPath                = $entityLog
        SqlTable               = if ($rows.Count -gt 0) { "$SqlSchema.$tableName" } else { $null }
    }
}

# ============================================================
# PREREQUISITI
# ============================================================
if (-not (Get-Command Get-AzAccessToken -ErrorAction SilentlyContinue)) {
    throw 'Get-AzAccessToken non disponibile. Installa/importa il modulo Az.Accounts.'
}
if (-not (Get-Command Get-AzContext -ErrorAction SilentlyContinue)) {
    throw 'Get-AzContext non disponibile. Installa/importa il modulo Az.Accounts.'
}
if (-not (Get-Command Connect-AzAccount -ErrorAction SilentlyContinue)) {
    throw 'Connect-AzAccount non disponibile. Installa/importa il modulo Az.Accounts.'
}
if (-not (Get-Command Get-AzServiceBusQueue -ErrorAction SilentlyContinue)) {
    throw 'Get-AzServiceBusQueue non disponibile. Installa/importa il modulo Az.ServiceBus.'
}
if (-not (Get-Command Get-AzServiceBusTopic -ErrorAction SilentlyContinue)) {
    throw 'Get-AzServiceBusTopic non disponibile. Installa/importa il modulo Az.ServiceBus.'
}
if (-not (Get-Command Get-AzServiceBusSubscription -ErrorAction SilentlyContinue)) {
    throw 'Get-AzServiceBusSubscription non disponibile. Installa/importa il modulo Az.ServiceBus.'
}

# ============================================================
# MAIN
# ============================================================
if (!(Test-Path $BasePath)) {
    New-Item -ItemType Directory -Path $BasePath -Force | Out-Null
}
if (Test-Path $GlobalLog) {
    Remove-Item $GlobalLog -Force
}

Write-Log '============================================================'
Write-Log 'CHECK + EXPORT DLQ (Queue e Topic/Subscription)'
Write-Log '============================================================'
Write-Log "Namespace       : $Namespace"
Write-Log "Resource Group  : $ResourceGroup"
Write-Log "Output report   : $OutputCsv"
Write-Log "BasePath        : $BasePath"
Write-Log ("Target entities : {0}" -f ($TargetEntities -join ', '))

Ensure-AzureLogin

$Results = New-Object System.Collections.Generic.List[object]
$http = [System.Net.Http.HttpClient]::new()
$http.Timeout = [TimeSpan]::FromMinutes(10)

try {
    foreach ($entity in $TargetEntities) {
        Write-Log "------------------------------------------------------------"
        Write-Log "Analisi entity: $entity"

        # 1) QUEUE con stesso nome dell'entity
        $queueInfo = Get-QueueDlqCount -QueueName $entity
        if ($queueInfo -and $queueInfo.Exists) {
            Write-Log "Queue trovata: $entity - DeadLetterMessageCount = $($queueInfo.DeadLetterMessageCount)"
            if ($queueInfo.DeadLetterMessageCount -gt 0) {
                $exportResult = Export-QueueDlq -QueueName $entity -DeadLetterMessageCount $queueInfo.DeadLetterMessageCount -HttpClient $http
                $Results.Add($exportResult)
            }
        }
        else {
            Write-Log "Queue non trovata: $entity" 'WARN'
        }

        # 2) TOPIC con stesso nome dell'entity + tutte le sue subscription
        $topicSubs = @(Get-TopicSubscriptionsDlqCount -TopicName $entity)
        if ($topicSubs.Count -eq 0) {
            Write-Log "Topic non trovato oppure nessuna subscription trovata: $entity" 'WARN'
        }
        else {
            foreach ($subInfo in $topicSubs) {
                Write-Log "Topic/Subscription trovata: $($subInfo.TopicName) / $($subInfo.SubscriptionName) - DeadLetterMessageCount = $($subInfo.DeadLetterMessageCount)"
                if ($subInfo.DeadLetterMessageCount -gt 0) {
                    $exportResult = Export-TopicSubscriptionDlq -TopicName $subInfo.TopicName -SubscriptionName $subInfo.SubscriptionName -DeadLetterMessageCount $subInfo.DeadLetterMessageCount -HttpClient $http
                    $Results.Add($exportResult)
                }
            }
        }
    }
}
catch {
    Write-Log "Errore durante elaborazione: $($_.Exception.Message)" 'ERROR'
    throw
}
finally {
    $http.Dispose()
}

Write-Log '============================================================'
Write-Log 'RISULTATI DLQ ATTIVE'
Write-Log '============================================================'

if ($Results.Count -eq 0) {
    Write-Log 'Nessuna DLQ trovata per le entity target'
}
else {
    $Results | Sort-Object EntityType, EntityName, SubscriptionName | ForEach-Object {
        if ($_.EntityType -eq 'Queue') {
            Write-Log ("QUEUE | {0} | DLQ={1} | Exported={2}" -f $_.EntityName, $_.DeadLetterMessageCount, $_.ExportedRows)
        }
        else {
            Write-Log ("TOPIC/SUB | {0} / {1} | DLQ={2} | Exported={3}" -f $_.EntityName, $_.SubscriptionName, $_.DeadLetterMessageCount, $_.ExportedRows)
        }
    }

    $Results | Export-Csv -Path $OutputCsv -NoTypeInformation -Delimiter ';' -Encoding UTF8
    Write-Log "CSV report esportato in: $OutputCsv"
}
