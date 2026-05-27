# ============================================
# CONFIG
# ============================================
$Namespace        = "YOUR_NAMESPACE"
$TopicName        = "YOUR_TOPIC"
$SubscriptionName = "YOUR_SUBSCRIPTION"

$TenantId     = "YOUR_TENANT_ID"
$ClientId     = "YOUR_APP_ID"
$ClientSecret = "YOUR_SECRET"

$DllFolder = "D:\PowerShell\Azure.Queue.Topic\libs"
$OutputCsv = "D:\PowerShell\DLQ_${TopicName}_${SubscriptionName}.csv"

$MaxMessages = 1000
$BatchSize   = 100

# ============================================
# LOAD DLL
# ============================================
Add-Type -Path (Join-Path $DllFolder "Azure.Core.dll")
Add-Type -Path (Join-Path $DllFolder "Azure.Identity.dll")
Add-Type -Path (Join-Path $DllFolder "Azure.Messaging.ServiceBus.dll")

# ============================================
# AUTH
# ============================================
$credential = [Azure.Identity.ClientSecretCredential]::new(
    $TenantId,
    $ClientId,
    $ClientSecret
)

$fqdn = "$Namespace.servicebus.windows.net"

$client = [Azure.Messaging.ServiceBus.ServiceBusClient]::new(
    $fqdn,
    $credential
)

# ============================================
# RECEIVER DLQ (SUBSCRIPTION)
# ============================================
$options = [Azure.Messaging.ServiceBus.ServiceBusReceiverOptions]::new()
$options.SubQueue = [Azure.Messaging.ServiceBus.SubQueue]::DeadLetter

$receiver = $client.CreateReceiver(
    $TopicName,
    $SubscriptionName,
    $options
)

# ============================================
# READ MESSAGES (PEEK - NO DELETE)
# ============================================
$results = @()
$fromSequenceNumber = $null
$totalRead = 0

while ($totalRead -lt $MaxMessages) {

    if ($null -eq $fromSequenceNumber) {
        $batch = $receiver.PeekMessagesAsync($BatchSize).GetAwaiter().GetResult()
    }
    else {
        $batch = $receiver.PeekMessagesAsync($BatchSize, [int64]$fromSequenceNumber).GetAwaiter().GetResult()
    }

    if ($batch.Count -eq 0) {
        break
    }

    foreach ($msg in $batch) {

        $results += [PSCustomObject]@{
            MessageId                  = $msg.MessageId
            DeadLetterReason           = $msg.DeadLetterReason
            DeadLetterErrorDescription = $msg.DeadLetterErrorDescription
        }

        $fromSequenceNumber = [int64]$msg.SequenceNumber + 1
        $totalRead++

        if ($totalRead -ge $MaxMessages) { break }
    }
}

# ============================================
# EXPORT CSV
# ============================================
if ($results.Count -gt 0) {
    $results | Export-Csv -Path $OutputCsv -NoTypeInformation -Delimiter ";"
    Write-Host "CSV creato: $OutputCsv"
}
else {
    Write-Host "Nessun messaggio in DLQ"
}

# ============================================
# CLEANUP
# ============================================
$receiver.DisposeAsync().GetAwaiter().GetResult() | Out-Null
$client.DisposeAsync().GetAwaiter().GetResult() | Out-Null