
# Parametri
$excelFolder = "\\SWOL003D\GBilder\03_TEXTE\ExcelFiles"   # Sostituisci con il percorso corretto
$sqlServer = "localhost"                 # Nome del server SQL
$database = "FIGURINO"                    # Nome del database
$table = "tblDocumentImageExcel"               # Nome della tabella



# Connessione SQL
$connectionString = "Server=$sqlServer;Database=$database;Integrated Security=True;"
$connection = New-Object System.Data.SqlClient.SqlConnection($connectionString)

try {
    $connection.Open()
    Write-Host "Connessione al database riuscita."
} catch {
    Write-Error "Errore di connessione al database: $($_.Exception.Message)"
    exit
}

# Funzione per eseguire INSERT
function Insert-Row($articleNumber, $productName, $language, $longDescription, $shortDescription, $editorialBulletpoints, $fileName) {
    try {
        $query = @"
INSERT INTO $table (articleNumber, productName, language, longDescription, shortDescription, editorialBulletpoints, fileName)
VALUES (@articleNumber, @productName, @language, @longDescription, @shortDescription, @editorialBulletpoints, @fileName)
"@

        $command = $connection.CreateCommand()
        $command.CommandText = $query

        $command.Parameters.AddWithValue("@articleNumber", $articleNumber) | Out-Null
        $command.Parameters.AddWithValue("@productName", $productName) | Out-Null
        $command.Parameters.AddWithValue("@language", $language) | Out-Null
        $command.Parameters.AddWithValue("@longDescription", $longDescription) | Out-Null
        $command.Parameters.AddWithValue("@shortDescription", $shortDescription) | Out-Null
        $command.Parameters.AddWithValue("@editorialBulletpoints", $editorialBulletpoints) | Out-Null
        $command.Parameters.AddWithValue("@fileName", $fileName) | Out-Null

        $command.ExecuteNonQuery() | Out-Null
    } catch {
        Write-Error "Errore durante l'inserimento nel DB: $($_.Exception.Message)"
    }
}

# Carica COM Excel
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

# Leggi tutti i file Excel nella cartella
$excelFiles = Get-ChildItem -Path $excelFolder -Filter *.xlsx

foreach ($file in $excelFiles) {
    Write-Host "Elaboro file: $($file.FullName)"

    try {
        # Apri Excel
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $workbook = $excel.Workbooks.Open($file.FullName)
        $sheet = $workbook.Sheets.Item(1)  # Unico sheet

        $lastRow = $sheet.UsedRange.Rows.Count

        # Leggi righe e inserisci nel DB
        for ($row = 2; $row -le $lastRow; $row++) {
            $articleNumber = $sheet.Cells.Item($row, 1).Text   # Colonna A
            $productName = $sheet.Cells.Item($row, 2).Text     # Colonna B
            $language = $sheet.Cells.Item($row, 4).Text        # Colonna D
            $longDescription = $sheet.Cells.Item($row, 5).Text # Colonna E
            $shortDescription = $sheet.Cells.Item($row, 6).Text# Colonna F
            $editorialBulletpoints = $sheet.Cells.Item($row, 7).Text # Colonna G

            
            if ($articleNumber -ne "") {
                Insert-Row $articleNumber $productName $language $longDescription $shortDescription $editorialBulletpoints $file.Name
            }
            else
            {
            
                $articleNumber = $articleNumber -replace '\s+', ''
                $productName = $articleNumber -replace '\s+', ''
                $language = $articleNumber -replace '\s+', ''
                $longDescription= $articleNumber -replace '\s+', ''
                $shortDescription= $articleNumber -replace '\s+', ''
                $editorialBulletpoints= $articleNumber -replace '\s+', ''
                if ($articleNumber -eq "" -and $productName -eq "" -and $language -eq "" -and $longDescription -eq "" -and $shortDescription -eq "" -and $editorialBulletpoints -eq "")
                {
                    break
                }
            }
        }

        # Chiudi Excel
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
    } catch {
        Write-Error "Errore durante l'elaborazione del file $($file.Name): $($_.Exception.Message)"
    }
}

$connection.Close()
Write-Host "Importazione completata!"
