
# Carica l'assembly di WinSCP
Add-Type -Path "C:\Program Files (x86)\WinSCP\WinSCPnet.dll"

# Crea l'oggetto SessionOptions
$sessionOptions = New-Object WinSCP.SessionOptions
$sessionOptions.Protocol = [WinSCP.Protocol]::Ftp
$sessionOptions.HostName = "ftp2.ups.com"
$sessionOptions.UserName = "wuerthlogistics0925"
$sessionOptions.Password = "IuRLP0HGl2PdCi8AFkVrahWMQEy8tylj"
$sessionOptions.FtpSecure = [WinSCP.FtpSecure]::Explicit
$sessionOptions.FtpMode = [WinSCP.FtpMode]::Passive

# Avvia la sessione
$session = New-Object WinSCP.Session
try {
    $session.Open($sessionOptions)

    # Percorsi
    $remoteFolder = "/"  # <-- aggiorna con il percorso corretto
    $localFolder = "D:\POD\"  
    $localNetworkFolder = "\\SWOL003D\Groups\GSOF\POD\" # oppure una cartella di rete
    $localBackup = "D:\POD\backup\"     # Cartella FTP di destinazione per i file copiati

    # Opzioni di trasferimento
    $transferOptions = New-Object WinSCP.TransferOptions
    $transferOptions.TransferMode = [WinSCP.TransferMode]::Binary
    # Scarica tutti i file
    $transferResult = $session.GetFiles("$remoteFolder*", $localFolder, $false, $transferOptions)

    # Verifica il risultato
    $transferResult.Check()

    $transferNetworkResult = $session.GetFiles("$remoteFolder*", $localNetworkFolder, $false, $transferOptions)
    $transferNetworkResult.Check()


    $transferLocalBackupResult = $session.GetFiles("$remoteFolder*", $localBackup, $false, $transferOptions)
    $transferLocalBackupResult.Check()


    Write-Host "✅ Download completato con successo."

    
    

    # Cancella i file dal server FTP dopo il download
    foreach ($file in $transferResult.Transfers) {
        $remotePath = $remoteFolder + $file.FileName
        $session.RemoveFiles($remotePath)
    }


    Write-Host "✅ Cancellazione remota completata con successo."


}
catch {
    Write-Host "❌ Errore: $($_.Exception.Message)"
}
finally {
    $session.Dispose()
}
