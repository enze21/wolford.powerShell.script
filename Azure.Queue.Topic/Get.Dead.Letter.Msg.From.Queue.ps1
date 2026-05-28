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

#PRD
$config = $config.PRD

#TEST
#$config = $config.TEST

#$TenantId      = $config.TenantId
#$SubscriptionId = $config.SubscriptionId
#$ResourceGroup = $config.ResourceGroup
$Namespace     = $config.Namespace
#$AppId         = $config.AppId
#$Secret        = $config.Secret | ConvertTo-SecureString -AsPlainText -Force
$BasePath      = $config.BasePath

$DateStamp = Get-Date -Format "dd.MM.yyyy"

#$EnableCsvExport = $true

$QueueName = "myqueue"

$OutputCsv = "$BasePath\$DateStamp.deadletters_entraid_extended.csv"
$LogFile   = "$BasePath\$DateStamp.deadletters_entraid_extended.log.txt"

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
        [Parameter(Mandatory = $true)]$Headers,
        [Parameter(Mandatory = $true)][string[]]$Names
    )

    foreach ($name in $Names) {
        foreach ($key in $Headers.Keys) {
            if ($key -ieq $name) {
                $value = $Headers[$key]
                if ($value -is [System.Array]) {
                    return ($value -join ",")
                }
                return [string]$value
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

if (Test-Path $LogFile) {
    Remove-Item $LogFile -Force
}

Write-Log "Avvio export DLQ con Entra ID"
Write-Log "Namespace: $Namespace"
Write-Log "Queue    : $QueueName"
Write-Log "CSV      : $OutputCsv"
Write-Log "Log      : $LogFile"

# ------------------------------------------------------------
# Login Azure se necessario
# ------------------------------------------------------------

try {
    $ctx = Get-AzContext -ErrorAction SilentlyContinue
    if (-not $ctx) {
        Write-Log "Nessuna sessione Azure attiva. Avvio Connect-AzAccount..."
        Connect-AzAccount | Out-Null
    } else {
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

$dlqEntityPath = "$QueueName/`$deadletterqueue"
$baseUri       = "https://$Namespace.servicebus.windows.net/$dlqEntityPath"
$headUri       = "$baseUri/messages/head?timeout=$ReceiveTimeoutSeconds"

Write-Log "DLQ path REST: $dlqEntityPath"

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
        foreach ($lm in @($lockedMessages)) {
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
                            $renewBrokerRaw = Get-HeaderValue -Headers $renewResp.Headers -Names @("BrokerProperties")
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

        if ($status -eq 204) {
            Write-Log "Nessun altro messaggio disponibile in DLQ"
            break
        }

        if ($status -ne 201) {
            throw "Errore HTTP durante Peek-Lock. Status=$status Reason=$($resp.ReasonPhrase)"
        }

        $location  = Get-HeaderValue -Headers $resp.Headers -Names @("Location")
        $brokerRaw = Get-HeaderValue -Headers $resp.Headers -Names @("BrokerProperties")

        $broker = $null
        if ($brokerRaw) {
            try {
                $broker = $brokerRaw | ConvertFrom-Json
            }
            catch {
                Write-Log "BrokerProperties non parsabile come JSON" "WARN"
            }
        }

        $messageId = Get-CaseInsensitiveProperty -Object $broker -Names @("MessageId")
        $sequenceNumber = Get-CaseInsensitiveProperty -Object $broker -Names @("SequenceNumber")
        $enqueuedTimeUtc = Get-CaseInsensitiveProperty -Object $broker -Names @("EnqueuedTimeUtc")
        if (-not $enqueuedTimeUtc) {
            $enqueuedTimeUtc = Get-HeaderValue -Headers $resp.Headers -Names @("Date")
        }

        $lockedUntil = Get-CaseInsensitiveProperty -Object $broker -Names @("LockedUntil", "LockedUntilUtc")

        $deadLetterReason =
            (Get-CaseInsensitiveProperty -Object $broker -Names @("DeadLetterReason")) ??
            (Get-HeaderValue -Headers $resp.Headers -Names @("DeadLetterReason", "deadletterreason"))

        $deadLetterDescription =
            (Get-CaseInsensitiveProperty -Object $broker -Names @("DeadLetterErrorDescription", "DeadLetterDescription")) ??
            (Get-HeaderValue -Headers $resp.Headers -Names @("DeadLetterErrorDescription", "DeadLetterDescription", "deadlettererrordescription", "deadletterdescription"))

        if ([string]::IsNullOrWhiteSpace($messageId) -and $sequenceNumber) {
            $messageId = "SequenceNumber:$sequenceNumber"
        }

        if ([string]::IsNullOrWhiteSpace($messageId)) {
            $messageId = "UNKNOWN"
        }

        $row = [PSCustomObject]@{
            MessageId             = $messageId
            SequenceNumber        = $sequenceNumber
            EnqueuedTimeUtc       = $enqueuedTimeUtc
            DeadLetterReason      = $deadLetterReason
            DeadLetterDescription = $deadLetterDescription
            ReadAtUtc             = (Get-Date).ToUniversalTime().ToString("o")
        }

        $rows.Add($row)

        $lockedMessages.Add([PSCustomObject]@{
            MessageId      = $messageId
            Location       = $location
            LockedUntilUtc = $(if ($lockedUntil) { [datetime]$lockedUntil } else { $null })
        })

        Write-Log "Messaggio acquisito | MessageId=$messageId | SequenceNumber=$sequenceNumber | Reason=$deadLetterReason"
    }

    # Export CSV
    $rows | Export-Csv -Path $OutputCsv -NoTypeInformation -Encoding UTF8

    Write-Log "Export CSV completato"
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
        if ([string]::IsNullOrWhiteSpace($lm.Location)) {
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