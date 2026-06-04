

function Write-Log {
    param (
        [string]$Message,
        [string]$LogFile = "C:\Logs\script_log.txt",
        [string]$Level = "INFO"
    )

    # Crea la cartella se non esiste
    $logDir = Split-Path -Path $LogFile
    if (-not (Test-Path -Path $logDir)) {
        New-Item -ItemType Directory -Path $logDir -Force | Out-Null
    }

    # Formatta il messaggio di log
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "$timestamp [$Level] $Message"

    # Scrive nel file
    Add-Content -Path $LogFile -Value $logEntry
}



# Percorso della directory di rete
$sourcePath = "\\SWOL002D\GBilder\eCommerce\eCommerce\DATABASE IMAGES"

Write-Log "Start " "D:\PowerShell\Immagini\log.lista.media.mp4.TLG.txt" 

# Percorso del file Excel
$excelPath = "D:\PowerShell\Immagini\lista media mp4 TLG.xlsx"

# Avvia Excel e carica il file
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Open($excelPath)
$sheet = $workbook.Sheets.Item(1)

$subfolderName = $sheet.Name

#$destinationPath = Join-Path -Path $sourcePath -ChildPath $subfolderName
$destinationPath = "\\SWOL002D\GBilder\eCommerce\eCommerce\DATABASE IMAGES\IT\Videos TLG"

# Legge fino a 1000 righe da colonne B e C
$identificatori = @()
for ($row = 2; $row -le 1001; $row++) {

    $valore1 = $sheet.Cells.Item($row, 1).Text

    $valore2 = $sheet.Cells.Item($row, 2).Text
    $valore3 = $sheet.Cells.Item($row, 3).Text

    $valore4 = $sheet.Cells.Item($row, 4).Text

    $valore5 = $sheet.Cells.Item($row, 5).Text

    Write-Host "row: $row valore1: $valore1 valore2: $valore2 valore3: $valore3 valore4: $valore4 valore5: $valore5"
     
    
    if (($valore5 -eq 'Y') -and $valore2 -and $valore3) {
        $identificatori = "${valore2}_${valore3}_*.mp4"
        Write-Log "Added $valore2 $valore3" "D:\PowerShell\Immagini\log.lista.media.mp4.TLG.txt" 
    
            # Crea la sottocartella se non esiste
        if (-not (Test-Path -Path $destinationPath)) {
            New-Item -ItemType Directory -Path $destinationPath | Out-Null
            Write-Log "Created Folder " "D:\PowerShell\Immagini\log.lista.media.mp4.TLG.txt" 

         }

        
        # Verifica che la directory esista
        if (-Not (Test-Path $sourcePath)) {
            Write-Host "La directory non esiste: $directoryPath"
            exit
        }

        # Ottieni tutti i file nella directory
        $allFiles = Get-ChildItem -Path $sourcePath -File -Recurse

        # Filtra i file che contengono il pattern nel nome
        #$fileArray = $allFiles | Where-Object { $_.Name -like $identificatori -and $_.CreationTime -gt (Get-Date "2024-09-01") }
        $fileArray = $allFiles | Where-Object { $_.Name -like $identificatori }

        # Verifica se sono stati trovati file
        if ($fileArray.Count -eq 0) {
            Write-Host "Nessun file trovato con il pattern specificato."
            Write-Log "Nessun file trovato con il pattern specificato." "D:\PowerShell\Immagini\log.lista.media.mp4.TLG.txt" 
        } else {
            # Stampa i file trovati
            Write-Host "File trovati:"
            Write-Log "File trovati:" "D:\PowerShell\Immagini\log.lista.media.mp4.TLG.txt" 
            $fileArray | ForEach-Object { Write-Host $_.FullName }
        }
        
        # Copia i file nella sottocartella
        foreach ($file in $fileArray) {
            Copy-Item -Path $file.FullName -Destination $destinationPath -Force
            $fileTemp = $file.FullName
            Write-Log "Copiato :  $fileTemp " "D:\PowerShell\Immagini\log.lista.media.mp4.TLG.txt" 

        }


       

   
    }
 
}

# Chiude Excel
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()



Write-Log "Finish " "D:\PowerShell\Immagini\log.lista.media.mp4.TLG.txt" 

