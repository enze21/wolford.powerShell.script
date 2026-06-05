# ============================================================
# CONFIGURAZIONE - MODIFICA SOLO QUESTE VARIABILI
# ============================================================

# ============================================
# LOAD CONFIG
# ============================================
$configPath = "D:\PowerShell\Azure.Queue.Topic\config.json"

if (!(Test-Path $configPath)) {
    Write-Error "Config file not found: $configPath"
    exit 1
}

$config = Get-Content $configPath | ConvertFrom-Json

$Namespace = $config.TEST.Namespace
$BasePath  = $config.TEST.BasePath

# SQL Server on-prem
$SqlServer                = $config.TEST.SqlServer
$SqlDatabase              = if ($config.TEST.SqlDatabase) { $config.TEST.SqlDatabase } else { "WIL.Queue.Topic" }
$SqlSchema                = if ($config.TEST.SqlSchema) { $config.TEST.SqlSchema } else { "dbo" }
$SqlUseIntegratedSecurity = if ($null -ne $config.TEST.SqlUseIntegratedSecurity) { [bool]$config.TEST.SqlUseIntegratedSecurity } else { $true }
$SqlUser                  = $config.TEST.SqlUser
$SqlPassword              = $config.TEST.SqlPassword

$DateStamp = Get-Date -Format "dd.MM.yyyy"

# ============================================================
# VARIABILI TOPIC / SUBSCRIPTION
# ============================================================
$TopicName        = "barcodeprojectionupdated"
$SubscriptionName = "Barcode.EDS.Out.M3.Connector"

# Nome tabella SQL ESATTAMENTE come richiesto:
# <nome topic>.<nome subscription>
# Verrà poi quotato come [dbo].[topic.subscription]
$SqlTableName = "$TopicName.$SubscriptionName"

$OutputCsv = "$BasePath\$DateStamp.$TopicName.$SubscriptionName.csv"
$LogFile   = "$BasePath\$DateStamp.$TopicName.$SubscriptionName.log.txt"

$ReceiveTimeoutSeconds = 3
$RenewBufferSeconds    = 10

# ============================================================
# NON MODIFICARE SOTTO
# ============================================================

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ------------------------------------------------------------
# Funzioni di utilità
# ------------------------------------------------------------

function Write-Log {
    param(
        [Parameter(Mandatory = $true)][string]$Message,
        [Parameter(Mandatory = $false)][string]$Level = "INFO"
    )

    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    $line = "[$timestamp] [$Level] $Message"

    Write-Host $line
    Add-Content -Path $LogFile -Value $line
}

function Convert-SecureStringToPlainText {
    param(
        [Parameter(Mandatory = $true)]
        [System.Security.SecureString]$SecureString
    )

    $ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)
    }
    finally {
        if ($ptr -ne [IntPtr]::Zero) {
            [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)
        }
    }
}

function Get-ServiceBusTokenInfo {
    $tok = Get-AzAccessToken -ResourceUrl "https://servicebus.azure.net/"
    $plain = Convert-SecureStringToPlainText -SecureString $tok.Token

    [PSCustomObject]@{
        AccessToken = $plain
        ExpiresOn   = $tok.ExpiresOn
    }
}

function Get-ValidServiceBusToken {
    param(
        [Parameter(Mandatory = $false)]
        $CurrentTokenInfo
    )

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

    $allHeaders = @()

    if ($Response.Headers) {
        $allHeaders += $Response.Headers
    }

    if ($Response.Content -and $Response.Content.Headers) {
        $allHeaders += $Response.Content.Headers
    }

    foreach ($header in $allHeaders) {
        foreach ($wanted in $Names) {
            if ($header.Key -ieq $wanted) {
                if (($header.Value -is [System.Collections.IEnumerable]) -and -not ($header.Value -is [string])) {
                    return ($header.Value -join ",")
                }
                return [string]$header.Value
            }
        }
    }

    return $null
}

function Get-CaseInsensitiveProperty {
    param(
        [Parameter(Mandatory = $true)]$Object,
        [Parameter(Mandatory = $true)][string[]]$Names
    )

    if ($null -eq $Object) {
        return $null
    }

    foreach ($wanted in $Names) {
        foreach ($prop in $Object.PSObject.Properties) {
            if ($prop.Name -ieq $wanted) {
                return $prop.Value
            }
        }
    }

    return $null
}

# ------------------------------------------------------------
# Funzioni Service Bus REST
# ------------------------------------------------------------

function Invoke-ServiceBusRequest {
    param(
        [Parameter(Mandatory = $true)][System.Net.Http.HttpClient]$HttpClient,
        [Parameter(Mandatory = $true)][string]$Method,
        [Parameter(Mandatory = $true)][string]$Uri,
        [Parameter(Mandatory = $true)][string]$AccessToken
    )

    $request = [System.Net.Http.HttpRequestMessage]::new(
        [System.Net.Http.HttpMethod]::new($Method),
        $Uri
    )

    $request.Headers.TryAddWithoutValidation("Authorization", "Bearer $AccessToken") | Out-Null
    $request.Content = [System.Net.Http.StringContent]::new("")

    $response = $HttpClient.SendAsync($request).GetAwaiter().GetResult()
    return $response
}

# ------------------------------------------------------------
# Funzioni di utilità SQL
# ------------------------------------------------------------

function Quote-SqlIdentifier {
    param(
        [Parameter(Mandatory = $true)][string]$Name
    )

    return "[" + $Name.Replace("]", "]]") + "]"
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
        return "Server=$Server;Database=$Database;Integrated Security=True;TrustServerCertificate=True;Encrypt=False;Application Name=DLQExportTopicSubscription;"
    }
    else {
        return "Server=$Server;Database=$Database;User ID=$User;Password=$Password;TrustServerCertificate=True;Encrypt=False;Application Name=DLQExportTopicSubscription;"
    }
}

function Convert-ToNullableDateTime {
    param(
        [Parameter(Mandatory = $false)]$Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return [DBNull]::Value
    }

    try {
        return [datetime]$Value
    }
    catch {
        return [DBNull]::Value
    }
}

function Convert-ToNullableInt64 {
    param(
        [Parameter(Mandatory = $false)]$Value
    )

    if ($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)) {
        return [DBNull]::Value
    }

    try {
        return [int64]$Value
    }
    catch {
        return [DBNull]::Value
    }
}

function Ensure-SqlDatabaseAndTable {
    param(
        [Parameter(Mandatory = $true)][string]$Server,
        [Parameter(Mandatory = $true)][string]$DatabaseName,
        [Parameter(Mandatory = $true)][string]$SchemaName,
        [Parameter(Mandatory = $true)][string]$TableName,
        [Parameter(Mandatory = $true)][bool]$UseIntegratedSecurity,
        [Parameter(Mandatory = $false)][string]$User,
        [Parameter(Mandatory = $false)][string]$Password
    )

    Add-Type -AssemblyName System.Data

    # 1. Crea DB se non esiste
    $masterConnString = Get-SqlConnectionString `
        -Server $Server `
        -Database "master" `
        -UseIntegratedSecurity $UseIntegratedSecurity `
        -User $User `
        -Password $Password

    $masterConn = New-Object System.Data.SqlClient.SqlConnection($masterConnString)
    $masterConn.Open()

    try {
        $cmd = $masterConn.CreateCommand()
        $cmd.CommandText = @"
IF DB_ID(@dbName) IS NULL
BEGIN
    DECLARE @sql nvarchar(max);
    SET @sql = N'CREATE DATABASE ' + QUOTENAME(@dbName);
    EXEC(@sql);
END
"@
        [void]$cmd.Parameters.Add("@dbName", [System.Data.SqlDbType]::NVarChar, 128)
        $cmd.Parameters["@dbName"].Value = $DatabaseName
        [void]$cmd.ExecuteNonQuery()
    }
    finally {
        $masterConn.Close()
        $masterConn.Dispose()
    }

    # 2. Crea tabella se non esiste
    $dbConnString = Get-SqlConnectionString `
        -Server $Server `
        -Database $DatabaseName `
        -UseIntegratedSecurity $UseIntegratedSecurity `
        -User $User `
        -Password $Password

    $dbConn = New-Object System.Data.SqlClient.SqlConnection($dbConnString)
    $dbConn.Open()

    try {
        $schemaQuoted = Quote-SqlIdentifier -Name $SchemaName
        $tableQuoted  = Quote-SqlIdentifier -Name $TableName
        $fullTable    = "$schemaQuoted.$tableQuoted"

        $cmd = $dbConn.CreateCommand()
        $cmd.CommandText = @"
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
        [void]$cmd.Parameters.Add("@tableName", [System.Data.SqlDbType]::NVarChar, 128)
        [void]$cmd.Parameters.Add("@schemaName", [System.Data.SqlDbType]::NVarChar, 128)
        $cmd.Parameters["@tableName"].Value = $TableName
        $cmd.Parameters["@schemaName"].Value = $SchemaName
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
        [Parameter(Mandatory = $true)][bool]$UseIntegratedSecurity,
        [Parameter(Mandatory = $false)][string]$User,
        [Parameter(Mandatory = $false)][string]$Password
    )

    Add-Type -AssemblyName System.Data

    $connString = Get-SqlConnectionString `
        -Server $Server `
        -Database $DatabaseName `
        -UseIntegratedSecurity $UseIntegratedSecurity `
        -User $User `
        -Password $Password

    $conn = New-Object System.Data.SqlClient.SqlConnection($connString)
    $conn.Open()

    try {
        $schemaQuoted = Quote-SqlIdentifier -Name $SchemaName
        $tableQuoted  = Quote-SqlIdentifier -Name $TableName
        $fullTable    = "$schemaQuoted.$tableQuoted"

        $insertSql = @"
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

        foreach ($r in $Rows) {
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = $insertSql

            [void]$cmd.Parameters.Add("@TopicName", [System.Data.SqlDbType]::NVarChar, 255)
            [void]$cmd.Parameters.Add("@SubscriptionName", [System.Data.SqlDbType]::NVarChar, 255)
            [void]$cmd.Parameters.Add("@MessageId", [System.Data.SqlDbType]::NVarChar, 255)
            [void]$cmd.Parameters.Add("@SequenceNumber", [System.Data.SqlDbType]::BigInt)
            [void]$cmd.Parameters.Add("@EnqueuedTimeUtc", [System.Data.SqlDbType]::DateTime2)
            [void]$cmd.Parameters.Add("@DeadLetterReason", [System.Data.SqlDbType]::NVarChar, 1024)
            [void]$cmd.Parameters.Add("@DeadLetterDescription", [System.Data.SqlDbType]::NVarChar, -1)
            [void]$cmd.Parameters.Add("@MessageBody", [System.Data.SqlDbType]::NVarChar, -1)
            [void]$cmd.Parameters.Add("@ReadAtUtc", [System.Data.SqlDbType]::DateTime2)

            $cmd.Parameters["@TopicName"].Value = [string]$r.TopicName
            $cmd.Parameters["@SubscriptionName"].Value = [string]$r.SubscriptionName
            $cmd.Parameters["@MessageId"].Value = if ([string]::IsNullOrWhiteSpace([string]$r.MessageId)) { "UNKNOWN" } else { [string]$r.MessageId }
            $cmd.Parameters["@SequenceNumber"].Value = Convert-ToNullableInt64 $r.SequenceNumber
            $cmd.Parameters["@EnqueuedTimeUtc"].Value = Convert-ToNullableDateTime $r.EnqueuedTimeUtc
            $cmd.Parameters["@DeadLetterReason"].Value = if ([string]::IsNullOrWhiteSpace([string]$r.DeadLetterReason)) { [DBNull]::Value } else { [string]$r.DeadLetterReason }
            $cmd.Parameters["@DeadLetterDescription"].Value = if ([string]::IsNullOrWhiteSpace([string]$r.DeadLetterDescription)) { [DBNull]::Value } else { [string]$r.DeadLetterDescription }
            $cmd.Parameters["@MessageBody"].Value = if ([string]::IsNullOrWhiteSpace([string]$r.MessageBody)) { [DBNull]::Value } else { [string]$r.MessageBody }
            $cmd.Parameters["@ReadAtUtc"].Value = Convert-ToNullableDateTime $r.ReadAtUtc

            [void]$cmd.ExecuteNonQuery()
            $cmd.Dispose()
        }
    }
    finally {
        $conn.Close()
        $conn.Dispose()
    }
}

# ------------------------------------------------------------
# Verifica prerequisiti Az.Accounts
# ------------------------------------------------------------

if (-not (Get-Command Get-AzAccessToken -ErrorAction SilentlyContinue)) {
    throw "Get-AzAccessToken non disponibile. Installa/importa il modulo Az.Accounts."
}

if (-not (Get-Command Get-AzContext -ErrorAction SilentlyContinue)) {
    throw "Get-AzContext non disponibile. Installa/importa il modulo Az.Accounts."
}

if (-not (Get-Command Connect-AzAccount -ErrorAction SilentlyContinue)) {
    throw "Connect-AzAccount non disponibile. Installa/importa il modulo Az.Accounts."
}

# ------------------------------------------------------------
# Inizializzazione log
# ------------------------------------------------------------

if (!(Test-Path $BasePath)) {
    New-Item -ItemType Directory -Path $BasePath -Force | Out-Null
}

if (Test-Path $LogFile) {
    Remove-Item $LogFile -Force
}

Write-Log "Avvio export DLQ TOPIC/SUBSCRIPTION con Entra ID"
Write-Log "Namespace       : $Namespace"
Write-Log "Topic           : $TopicName"
Write-Log "Subscription    : $SubscriptionName"
Write-Log "CSV             : $OutputCsv"
Write-Log "Log             : $LogFile"

Write-Log "SQL Server      : $SqlServer"
Write-Log "SQL DB          : $SqlDatabase"
Write-Log "SQL Schema      : $SqlSchema"
Write-Log "SQL Table       : $SqlTableName"

# ------------------------------------------------------------
# Login Azure se necessario
# ------------------------------------------------------------

try {
    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        Write-Log "Nessuna sessione Azure attiva. Avvio Connect-AzAccount..."
        Connect-AzAccount | Out-Null
    }
    else {
        Write-Log "Sessione Azure già attiva: $($ctx.Account.Id)"
    }
}
catch {
    Write-Log "Errore nel recupero contesto Azure. Avvio Connect-AzAccount..." "WARN"
    Connect-AzAccount | Out-Null
}

# ------------------------------------------------------------
# Setup runtime
# ------------------------------------------------------------

$dlqEntityPath = "$TopicName/Subscriptions/$SubscriptionName/`$deadletterqueue"
$baseUri       = "https://$Namespace.servicebus.windows.net/$dlqEntityPath"
$headUri       = "$baseUri/messages/head?timeout=$ReceiveTimeoutSeconds"

Write-Log "DLQ path REST   : $dlqEntityPath"

$http = [System.Net.Http.HttpClient]::new()
$http.Timeout = [TimeSpan]::FromMinutes(10)

$rows = New-Object System.Collections.Generic.List[object]
$lockedMessages = New-Object System.Collections.Generic.List[object]

$tokenInfo = $null

# ------------------------------------------------------------
# Main
# ------------------------------------------------------------

try {
    while ($true) {

        $tokenInfo = Get-ValidServiceBusToken -CurrentTokenInfo $tokenInfo

        # Rinnovo lock dei messaggi già letti se stanno per scadere
        $now = Get-Date
        foreach ($lm in $lockedMessages.ToArray()) {
            if ($null -ne $lm.LockedUntilUtc) {
                $remaining = $lm.LockedUntilUtc - $now

                if ($remaining.TotalSeconds -le $RenewBufferSeconds) {
                    try {
                        $tokenInfo = Get-ValidServiceBusToken -CurrentTokenInfo $tokenInfo
                        $renewResp = Invoke-ServiceBusRequest `
                            -HttpClient $http `
                            -Method "POST" `
                            -Uri $lm.Location `
                            -AccessToken $tokenInfo.AccessToken

                        if ([int]$renewResp.StatusCode -eq 200) {
                            $renewBrokerRaw = Get-HeaderValue -Response $renewResp -Names @("BrokerProperties")
                            if ($renewBrokerRaw) {
                                $renewBroker = $renewBrokerRaw | ConvertFrom-Json
                                $newLockedUntil = Get-CaseInsensitiveProperty -Object $renewBroker -Names @("LockedUntil", "LockedUntilUtc")
                                if ($newLockedUntil) {
                                    $lm.LockedUntilUtc = [datetime]$newLockedUntil
                                    Write-Log "Lock rinnovato per messaggio $($lm.MessageId)"
                                }
                            }
                        }
                    }
                    catch {
                        Write-Log "Rinnovo lock fallito per $($lm.MessageId): $($_.Exception.Message)" "WARN"
                    }
                }
            }
        }

        $resp = Invoke-ServiceBusRequest `
            -HttpClient $http `
            -Method "POST" `
            -Uri $headUri `
            -AccessToken $tokenInfo.AccessToken

        $status = [int]$resp.StatusCode

        $body = $null
        try {
            if ($resp.Content) {
                $body = [System.Text.Encoding]::UTF8.GetString(
                    $resp.Content.ReadAsByteArrayAsync().GetAwaiter().GetResult()
                )
                $body = $body -replace "`r|`n", " "
            }
        }
        catch {
            Write-Log "Errore lettura body: $($_.Exception.Message)" "WARN"
        }

        if ($status -eq 204) {
            Write-Log "Nessun altro messaggio disponibile in DLQ"
            break
        }

        if ($status -ne 201) {
            throw "Errore HTTP durante Peek-Lock. Status=$status Reason=$($resp.ReasonPhrase)"
        }

        $location  = Get-HeaderValue -Response $resp -Names @("Location")
        $brokerRaw = Get-HeaderValue -Response $resp -Names @("BrokerProperties")

        $broker = $null
        if ($brokerRaw) {
            try {
                $broker = $brokerRaw | ConvertFrom-Json
            }
            catch {
                Write-Log "BrokerProperties non parsabile come JSON" "WARN"
            }
        }

        $messageId       = Get-CaseInsensitiveProperty -Object $broker -Names @("MessageId")
        $sequenceNumber  = Get-CaseInsensitiveProperty -Object $broker -Names @("SequenceNumber")
        $enqueuedTimeUtc = Get-CaseInsensitiveProperty -Object $broker -Names @("EnqueuedTimeUtc")

        if (-not $enqueuedTimeUtc) {
            $enqueuedTimeUtc = Get-HeaderValue -Response $resp -Names @("Date")
        }

        $lockedUntil = Get-CaseInsensitiveProperty -Object $broker -Names @("LockedUntil", "LockedUntilUtc")

        $deadLetterReason = Get-CaseInsensitiveProperty -Object $broker -Names @("DeadLetterReason")
        if ([string]::IsNullOrWhiteSpace([string]$deadLetterReason)) {
            $deadLetterReason = Get-HeaderValue -Response $resp -Names @("DeadLetterReason", "deadletterreason")
        }

        $deadLetterDescription = Get-CaseInsensitiveProperty -Object $broker -Names @("DeadLetterErrorDescription", "DeadLetterDescription")
        if ([string]::IsNullOrWhiteSpace([string]$deadLetterDescription)) {
            $deadLetterDescription = Get-HeaderValue -Response $resp -Names @("DeadLetterErrorDescription", "DeadLetterDescription", "deadlettererrordescription", "deadletterdescription")
        }

        if ([string]::IsNullOrWhiteSpace([string]$messageId) -and $sequenceNumber) {
            $messageId = "SequenceNumber:$sequenceNumber"
        }

        if ([string]::IsNullOrWhiteSpace([string]$messageId)) {
            $messageId = "UNKNOWN"
        }

        $row = [PSCustomObject]@{
            TopicName             = $TopicName
            SubscriptionName      = $SubscriptionName
            MessageId             = $messageId
            SequenceNumber        = $sequenceNumber
            EnqueuedTimeUtc       = $enqueuedTimeUtc
            DeadLetterReason      = $deadLetterReason
            DeadLetterDescription = $deadLetterDescription
            MessageBody           = $body
            ReadAtUtc             = (Get-Date).ToUniversalTime().ToString("o")
        }

        $rows.Add($row)

        $lockedMessages.Add([PSCustomObject]@{
            MessageId      = $messageId
            Location       = $location
            LockedUntilUtc = if ($lockedUntil) { [datetime]$lockedUntil } else { $null }
        })

        Write-Log "Messaggio acquisito | Topic=$TopicName | Subscription=$SubscriptionName | MessageId=$messageId | SequenceNumber=$sequenceNumber | Reason=$deadLetterReason"
    }

    # Export CSV
    $rows | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8
    Write-Log "Export CSV completato"

    # Ensure DB + table
    Ensure-SqlDatabaseAndTable `
        -Server $SqlServer `
        -DatabaseName $SqlDatabase `
        -SchemaName $SqlSchema `
        -TableName $SqlTableName `
        -UseIntegratedSecurity $SqlUseIntegratedSecurity `
        -User $SqlUser `
        -Password $SqlPassword

    Write-Log "Verifica DB/tabella SQL completata"

    # Export SQL (append)
    if ($rows.Count -gt 0) {
        Export-RowsToSql `
            -Rows $rows `
            -Server $SqlServer `
            -DatabaseName $SqlDatabase `
            -SchemaName $SqlSchema `
            -TableName $SqlTableName `
            -UseIntegratedSecurity $SqlUseIntegratedSecurity `
            -User $SqlUser `
            -Password $SqlPassword

        Write-Log "Export SQL completato"
    }
    else {
        Write-Log "Nessuna riga da inserire in SQL"
    }

    Write-Log "Totale messaggi esportati: $($rows.Count)"
}
catch {
    Write-Log "ERRORE BLOCCANTE: $($_.Exception.Message)" "ERROR"
    throw
}
finally {
    # Unlock finale
    Write-Log "Avvio unlock finale dei messaggi lockati"

    foreach ($lm in $lockedMessages) {
        if ([string]::IsNullOrWhiteSpace([string]$lm.Location)) {
            Write-Log "Unlock saltato per $($lm.MessageId): Location assente" "WARN"
            continue
        }

        try {
            $tokenInfo = Get-ValidServiceBusToken -CurrentTokenInfo $tokenInfo

            $unlockResp = Invoke-ServiceBusRequest `
                -HttpClient $http `
                -Method "PUT" `
                -Uri $lm.Location `
                -AccessToken $tokenInfo.AccessToken

            $unlockCode = [int]$unlockResp.StatusCode
            if ($unlockCode -eq 200) {
                Write-Log "Unlock OK per $($lm.MessageId)"
            }
            else {
                Write-Log "Unlock fallito per $($lm.MessageId). HTTP $unlockCode" "WARN"
            }
        }
        catch {
            Write-Log "Unlock eccezione per $($lm.MessageId): $($_.Exception.Message)" "WARN"
        }
    }

    if ($http) {
        $http.Dispose()
    }

    Write-Log "Fine script"
}