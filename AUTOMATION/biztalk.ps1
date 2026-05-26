# --- CONFIGURAZIONE ---
$serverName = "SBIZDEV1"
$dbName = "BizTalkMgmtDb"
$tableName = "adm_ReceiveLocation"
$targetLocationName = "rl.Wolford.Test"
$user = "userBT"
# Inserisci qui la password che mi manderai via mail
$password = "UserBT01@." 
$webhookUrl = "https://ai.nubble.it/webhook/biztalk-receive-locations"

# --- CONNESSIONE SQL ---
$connString = "Server=$serverName;Database=$dbName;User Id=$user;Password=$password;"
$query = "SELECT [Disabled] FROM [dbo].[$tableName] WHERE [Name] = '$targetLocationName'"

try {
    # Creazione connessione
    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $connString
    $connection.Open()

    # Esecuzione Query
    $command = $connection.CreateCommand()
    $command.CommandText = $query
    $result = $command.ExecuteScalar() # Recupera il singolo valore [Disabled]

    $connection.Close()

    # --- LOGICA DI CONTROLLO ---
    # Se il risultato è null, la location non esiste
    if ($null -eq $result) {
        Write-Host "Errore: Receive Location non trovata."
        $statusMessage = "NotFound"
        $statusId = -999
    }
    else {
        # Interpretazione risultato
        # 0 = Enabled (Attiva)
        # -1 = Disabled (Disabilitata)
        if ($result -eq 0) {
            $statusMessage = "Enabled"
            $statusId = 0
            Write-Host "Stato: Attivo (Enabled)"
        }
        elseif ($result -eq -1) {
            $statusMessage = "Disabled"
            $statusId = -1
            Write-Host "Stato: Disabilitato (Disabled)"
        }
        else {
            $statusMessage = "Unknown"
            $statusId = $result
        }
    }

    # --- INVIO AL WEBHOOK (Simile a cURL) ---
    
    # Creiamo il payload JSON
    $payload = @{
        env = "BizTalk-SBIZDEV1"
        location = $targetLocationName
        status = $statusMessage
        statusCode = $statusId
        timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    } | ConvertTo-Json

    # Inviamo la richiesta POST
    Write-Host "Invio dati al Webhook..."
    $response = Invoke-RestMethod -Uri $webhookUrl -Method Post -ContentType "application/json" -Body $payload
    
    Write-Host "Webhook risponde: " $response

}
catch {
    Write-Error "Si è verificato un errore: $_"
    # Opzionale: Inviare un alert al webhook anche in caso di crash dello script
}

#Read-Host -Prompt "Premi INVIO per chiudere questa finestra..."