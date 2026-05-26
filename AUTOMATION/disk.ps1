param(
    # URL del webhook n8n (usa /webhook-test/... per i test, /webhook/... in produzione)
    [string]$WebHookUrl = "https://ai.nubble.it/webhook/disk-report",

    # Soglia percentuale di spazio libero sotto la quale considerare il disco 'critico'
    [int]$WarnThreshold = 15,

    # Computer da analizzare (default: il computer locale)
    [string]$ComputerName = $env:COMPUTERNAME
)

Write-Host "Raccolta informazioni dischi fisici da [$ComputerName]..." -ForegroundColor Cyan

try {
    # Dischi fisici: DriveType = 3
    $logicalDisks = Get-CimInstance -ClassName Win32_LogicalDisk -ComputerName $ComputerName -Filter "DriveType = 3"
}
catch {
    Write-Host "Errore nel recupero delle informazioni dei dischi: $($_.Exception.Message)" -ForegroundColor Red
    Read-Host "`nPremi INVIO per chiudere"
    exit 1
}

if (-not $logicalDisks) {
    Write-Host "Nessun disco fisico trovato su [$ComputerName]." -ForegroundColor Yellow
}

# Costruisco la lista degli items
$items = @()

foreach ($d in $logicalDisks) {
    $size = [double]($d.Size)
    $free = [double]($d.FreeSpace)

    if ($size -le 0) {
        $percentFree = 0
    } else {
        $percentFree = [math]::Round(($free / $size) * 100, 2)
    }

    $obj = [pscustomobject]@{
        Volume        = $d.DeviceID           # es. "C:"
        DriveType     = "Fixed"               # fisso perché abbiamo filtrato DriveType=3
        SizeGB        = [math]::Round($size / 1GB, 2)
        FreeGB        = [math]::Round($free / 1GB, 2)
        PercentFree   = $percentFree
        BelowThreshold= ($percentFree -lt $WarnThreshold)
    }

    $items += $obj
}

# Corpo del messaggio per n8n
$body = @{
    source    = "disk"
    computer  = $ComputerName
    timestamp = (Get-Date).ToString("o")   # ISO 8601
    items     = $items
}

# Converto in JSON
$json = $body | ConvertTo-Json -Depth 5

Write-Host "`n--- REQUEST BODY (REAL DISKS) ---" -ForegroundColor Yellow
Write-Host $json
Write-Host "----------------------------------`n"

# Invio al webhook n8n
Write-Host "Invio payload a: $WebHookUrl" -ForegroundColor Cyan

try {
    $resp = Invoke-RestMethod -Method Post -Uri $WebHookUrl -Body $json -ContentType "application/json"
    Write-Host "`nResponse:" -ForegroundColor Green
    $resp | Out-String | Write-Host
}
catch {
    Write-Host "`nERRORE durante la richiesta al webhook:" -ForegroundColor Red
    Write-Host $_.Exception.Message
}

#Read-Host "`nPremi INVIO per chiudere"
