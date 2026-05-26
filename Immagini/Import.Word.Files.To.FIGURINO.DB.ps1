
# Parametri di connessione SQL Server
$server = "localhost"
$database = "FIGURINO"
$table = "tblDocumentImageWord"

# Percorso dei file Word
$folderPath = "\\SWOL003D\GBilder\03_TEXTE\Cumulus\WordFiles"

# File di log
$logFile = "D:\PowerShell\log_importazione.txt"

# Connection string con autenticazione Windows
$connectionString = "Server=$server;Database=$database;Integrated Security=True"

# Carica l'assembly SQL Server
Add-Type -AssemblyName "System.Data"

# Funzione per scrivere nel log
function Write-Log {
    param ([string]$message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Add-Content -Path $logFile -Value "$timestamp - $message"
}

# Inizio script
Write-Log "--- INIZIO IMPORTAZIONE DOCUMENTI ---"

try {
    $connection = New-Object System.Data.SqlClient.SqlConnection
    $connection.ConnectionString = $connectionString
    $connection.Open()
    Write-Log "Connessione al database riuscita."

    $word = New-Object -ComObject Word.Application
    $word.Visible = $false

    Get-ChildItem -Path $folderPath -Filter *.doc* | ForEach-Object {
        $filePath = $_.FullName
        $fileName = $_.Name

        try {
            $doc = $word.Documents.Open($filePath, $false, $true)
            $text = $doc.Content.Text.Trim()
            $doc.Close()

            $query = "INSERT INTO $table (fileName, contentText, Note, Description, Status) VALUES (@nome_file, @contenuto, '', '', 0)"
            $command = $connection.CreateCommand()
            $command.CommandText = $query

            $command.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@nome_file", $fileName)))
            $command.Parameters.Add((New-Object Data.SqlClient.SqlParameter("@contenuto", $text)))

            $command.ExecuteNonQuery()
            Write-Log "Importato: $fileName"
        }
        catch {
            Write-Log "ERRORE con file ${fileName}: $_"
        }
    }

    $word.Quit()
    $connection.Close()
    Write-Log "Importazione completata."
}
catch {
    Write-Log "ERRORE GENERALE: $_"
}

Write-Log "--- FINE SCRIPT ---"
