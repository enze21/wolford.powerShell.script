# =========================
# Test Webhook: Azure Report
# =========================

# ?? METTI QUI IL TUO WEBHOOK TEST DI AZURE
# es: https://ai.nubble.it/webhook-test/azure-report
$webhookUrl = "https://ai.nubble.it/webhook-test/azure-report"

Write-Host "Inviando payload di test ad Azure Report Webhook..." -ForegroundColor Cyan

# Oggetto JSON di test
$body = @{
    source    = "azure"
    report    = "activityLog"
    timestamp = "2025-11-17T21:00:00Z"
    items     = @(
        @{
            id       = "e1"
            severity = "Error"
            message  = "CPU spike"
            failed   = $true
        },
        @{
            id       = "e2"
            severity = "Warning"
            message  = "Latency high"
            failed   = $false
        },
        @{
            id       = "e3"
            severity = "Error"
            message  = "Out of memory"
            failed   = $true
        }
    )
}

# Converto in JSON
$json = $body | ConvertTo-Json -Depth 5

Write-Host "`n--- REQUEST BODY (Azure) ---" -ForegroundColor Yellow
Write-Host $json
Write-Host "-----------------------------`n"

# Invio
try {
    $resp = Invoke-RestMethod -Method Post -Uri $webhookUrl -Body $json -ContentType "application/json"
    Write-Host "Response:" -ForegroundColor Green
    $resp | Out-String | Write-Host
}
catch {
    Write-Host "`nERRORE durante la richiesta:" -ForegroundColor Red
    Write-Host $_.Exception.Message
}

Read-Host "`nPremi INVIO per chiudere"
