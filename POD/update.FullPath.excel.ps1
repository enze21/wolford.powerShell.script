
# Parametri di connessione
$server = "localhost"
$database = "FIGURINO"
$table = "tblDocumentImageExcel"
$pathRete = "\\SWOL003D\GBilder\03_TEXTE\Cumulus"

# Query per ottenere i nomi dei file
$querySelect = "SELECT id, fileName FROM $table WHERE filePath IS NULL"

# File di log
$logFile = "D:\PowerShell\log_importazione_update.txt"

# Connessione e lettura
$files = Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $querySelect

# Funzione per scrivere nel log
function Write-Log {
    param ([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "$timestamp - $message"
}

foreach ($file in $files) {
    $fileName = $file.FileName
    $id = $file.Id

    # Ricerca ricorsiva del file nel path di rete
    $foundFile = Get-ChildItem -Path $pathRete -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq $fileName }

    if ($foundFile) {
        $fullPath = $foundFile.FullName.Replace("'", "''")  # Escape per SQL

        # Aggiornamento della tabella
        $queryUpdate = "UPDATE $table SET filePath = '$fullPath' WHERE id = $id"
        Invoke-Sqlcmd -ServerInstance $server -Database $database -Query $queryUpdate
        Write-Log "Aggiornato: $fileName -> $fullPath"
    } else {
        Write-Log "File non trovato: $fileName"
    }
}
