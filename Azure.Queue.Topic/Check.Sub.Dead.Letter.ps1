
# ============================================
# LOAD CONFIG
# ============================================
$configPath = "D:\PowerShell\Azure.Queue.Topic\config.json"

if (!(Test-Path $configPath)) {
    Write-Error "Config file not found: $configPath"
    exit 1
}

$config = Get-Content $configPath | ConvertFrom-Json

$TenantId      = $config.TenantId
$SubscriptionId = $config.SubscriptionId
$ResourceGroup = $config.ResourceGroup
$Namespace     = $config.Namespace
$AppId         = $config.AppId
$Secret        = $config.Secret | ConvertTo-SecureString -AsPlainText -Force
$BasePath      = $config.BasePath

$DateStamp = Get-Date -Format "dd.MM.yyyy"
$OutputCsv = "$BasePath\$DateStamp.ServiceBus_DLQ_Report.csv"

$EnableCsvExport = $true

# ============================================
# LOG FUNCTION
# ============================================
function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Write-Host "$timestamp [$Level] - $Message"
}


# ============================================
# AZ LOGIN & CONTEXT (ROBUSTO)
# ============================================
Write-Log "Login su Azure Tenant $TenantId"

try {
    
    $AppId = $AppId
    $Secret = $Secret

    $Credential = New-Object System.Management.Automation.PSCredential ($AppId, $Secret)

    Connect-AzAccount `
        -ServicePrincipal `
        -Credential $Credential `
        -Tenant $TenantId `
        -Subscription $SubscriptionId


}
catch {
    Write-Log "Errore durante login/context: $_" "ERROR"
    exit 1
}



# ============================================
# MAIN
# ============================================
$Results = @()

try {

    # ============================================
    # TOPICS + SUBSCRIPTIONS
    # ============================================
    Write-Log "Recupero Topics dal namespace $Namespace"

    $topics = Get-AzServiceBusTopic `
        -ResourceGroupName $ResourceGroup `
        -NamespaceName  $Namespace

    foreach ($topic in $topics) {

        Write-Log "Analizzo topic: $($topic.Name)"

        $subscriptions = Get-AzServiceBusSubscription `
            -ResourceGroupName $ResourceGroup `
            -NamespaceName $Namespace `
            -TopicName $topic.Name

        foreach ($sub in $subscriptions) {

            $dlqCount = $sub.CountDetailDeadLetterMessageCount

            Write-Log "Sub: $($sub.Name) - DLQ: $dlqCount"

            if ($dlqCount -gt 0) {

                Write-Log ">>> DLQ trovata su Topic=$($topic.Name) Sub=$($sub.Name)" "WARNING"

                $Results += [PSCustomObject]@{
                    EntityType       = "Topic"
                    TopicName        = $topic.Name
                    SubscriptionName = $sub.Name
                    QueueName        = ""
                    DeadLetterCount  = $dlqCount
                }
            }
        }
    }

    # ============================================
    # QUEUES
    # ============================================
    Write-Log "Recupero Queues dal namespace $Namespace"

    $queues = Get-AzServiceBusQueue `
        -ResourceGroupName $ResourceGroup `
        -NamespaceName $Namespace

    foreach ($queue in $queues) {

        $dlqCount = $queue.CountDetailDeadLetterMessageCount

        #$dlqCount = $queue.CountDetails.DeadLetterMessageCount

        Write-Log "Queue: $($queue.Name) - DLQ: $dlqCount"

        if ($dlqCount -gt 0) {

            Write-Log ">>> DLQ trovata su Queue=$($queue.Name)" "WARNING"

            $Results += [PSCustomObject]@{
                EntityType       = "Queue"
                TopicName        = ""
                SubscriptionName = ""
                QueueName        = $queue.Name
                DeadLetterCount  = $dlqCount
            }
        }
    }

}
catch {
    Write-Log "Errore durante elaborazione: $_" "ERROR"
    exit 1
}

# ============================================
# RISULTATI
# ============================================
Write-Log "========================================="
Write-Log "RISULTATI DLQ ATTIVE"
Write-Log "========================================="

if ($Results.Count -eq 0) {
    Write-Log "Nessuna DLQ trovata"
}
else {

    # Topics distinti
    $TopicsWithDLQ = $Results |
        Where-Object { $_.EntityType -eq "Topic" } |
        Select-Object TopicName -Unique

    # Queues distinte
    $QueuesWithDLQ = $Results |
        Where-Object { $_.EntityType -eq "Queue" } |
        Select-Object QueueName -Unique

    Write-Log "--- Topics con DLQ ---"
    $TopicsWithDLQ | ForEach-Object {
        Write-Host " - $($_.TopicName)"
    }

    Write-Log "--- Queues con DLQ ---"
    $QueuesWithDLQ | ForEach-Object {
        Write-Host " - $($_.QueueName)"
    }

    Write-Log "Totale record DLQ trovati: $($Results.Count)"
}

# ============================================
# EXPORT CSV
# ============================================
$Results | Export-Csv -Path $OutputCsv -NoTypeInformation -Delimiter ";"

Write-Log "CSV esportato in: $OutputCsv"

if ($EnableCsvExport -and $Results.Count -gt 0) {

    $Results | Export-Csv -Path $OutputCsv -NoTypeInformation -Delimiter ";"

    Write-Log "CSV esportato in: $OutputCsv"
}
else {
    Write-Log "Nessun dato da esportare"
}