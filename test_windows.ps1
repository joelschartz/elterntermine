# ===== ElternTermine: Windows-Machbarkeitstest (PowerShell) =====
# Prueft, ob die PowerShell-Variante auf diesem PC ueberhaupt laufen kann
# (ohne Admin, ohne Installation). Es wird NICHTS gesendet oder gespeichert.

Write-Host "=================================================="
Write-Host "  ElternTermine - Windows-Test"
Write-Host "=================================================="
Write-Host ("PowerShell-Version : " + $PSVersionTable.PSVersion)
$mode = $ExecutionContext.SessionState.LanguageMode
Write-Host ("Sprachmodus        : " + $mode + "   (noetig: FullLanguage)")
Write-Host ""

$ok1 = $false; $ok2 = $false

# Test 1: Lokalen Port oeffnen (ohne Admin) - dafuer TcpListener (braucht keine Adminrechte)
try {
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 8765)
    $listener.Start()
    Start-Sleep -Milliseconds 100
    $listener.Stop()
    Write-Host "Test 1 (lokaler Mini-Server auf Port 8765): OK"
    $ok1 = $true
} catch {
    Write-Host ("Test 1 FEHLER: " + $_.Exception.Message)
}

# Test 2: Microsoft-Login erreichbar (Geraetecode anfordern - keine Anmeldung noetig)
try {
    $body = @{
        client_id = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
        scope     = "https://graph.microsoft.com/Mail.Send offline_access"
    }
    $r = Invoke-RestMethod -Method Post -TimeoutSec 30 `
        -Uri "https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode" -Body $body
    if ($r.user_code) {
        Write-Host ("Test 2 (Microsoft erreichbar): OK - Beispiel-Code " + $r.user_code)
        $ok2 = $true
    } else {
        Write-Host "Test 2: kein Code erhalten."
    }
} catch {
    Write-Host ("Test 2 FEHLER: " + $_.Exception.Message)
}

Write-Host ""
Write-Host "=================================================="
if ($ok1 -and $ok2 -and $mode -eq "FullLanguage") {
    Write-Host "  ERGEBNIS: ALLES OK - die PowerShell-Variante kann gebaut werden."
} else {
    Write-Host "  ERGEBNIS: Etwas ist gesperrt (siehe oben)."
    Write-Host "  Bitte die Zeilen Test 1 / Test 2 / Sprachmodus an Claude weitergeben."
}
Write-Host "=================================================="
Write-Host ""
Read-Host "Zum Schliessen Enter druecken"
