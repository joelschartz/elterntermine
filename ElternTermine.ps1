# =====================================================================
#  ElternTermine - lokaler Helfer fuer Windows (PowerShell)
#  Ersetzt server.py: liefert graph.html aus und erledigt Login (Geraetecode)
#  + Versand ueber Microsoft Graph. Kein Python, keine Installation, kein Admin.
#  Start ueber: ElternTermine-Start.bat
# =====================================================================

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$ClientId = "14d82eec-204b-4c2f-b7e8-296a70dab67e"   # Microsoft Graph Command Line Tools (oeffentlich)
$Tenant = "organizations"
$Scope  = "https://graph.microsoft.com/Mail.Send offline_access"
$Base   = "https://login.microsoftonline.com/$Tenant/oauth2/v2.0"
$Port   = 8765

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$TokenFile = Join-Path $ScriptDir "graph_token.json"
$LogFile   = Join-Path $ScriptDir "ElternTermine-log.txt"
$script:Pending = $null

function Log($msg) {
    $line = ((Get-Date -Format "HH:mm:ss") + "  " + $msg)
    Write-Host $line
    try { Add-Content -Path $LogFile -Value $line -Encoding UTF8 } catch {}
}

# Laedt beim Start die neueste graph.html aus dem GitHub-Repo und ersetzt die lokale Kopie.
# So bekommen alle Nutzer App-Updates ohne Neuinstallation. Schlaegt es fehl (offline /
# gesperrt), bleibt die vorhandene lokale Datei erhalten.
function Update-UiFromGitHub {
    $rawUrl = "https://raw.githubusercontent.com/joelschartz/elterntermine/main/graph.html"
    $target = Join-Path $ScriptDir "graph.html"
    try {
        Write-Host "  Suche nach Update ..."
        $resp = Invoke-WebRequest -Uri $rawUrl -UseBasicParsing -TimeoutSec 8
        $html = [string]$resp.Content
        if ($html.Contains("tplBodyEditor") -and $html.ToLower().Contains("<html")) {
            [IO.File]::WriteAllText($target, $html, (New-Object Text.UTF8Encoding($false)))
            Write-Host "  Oberflaeche auf neueste Version aktualisiert (GitHub)."
        } else {
            Write-Host "  (Update-Antwort unerwartet - nutze lokale Oberflaeche.)"
        }
    } catch {
        Write-Host "  (Offline / GitHub nicht erreichbar - nutze lokale Oberflaeche.)"
    }
}

# ----------------------------------------------------------- HTTP-Helfer
function PostForm($url, $form) {
    # Genau wie der funktionierende Test: Hashtable-Body, KEIN expliziter ContentType.
    Log ("PostForm -> " + $url)
    Log ("  client_id = " + $form["client_id"] + " | scope = " + $form["scope"])
    try {
        $r = Invoke-RestMethod -Method Post -Uri $url -Body $form -TimeoutSec 30
        if ($null -eq $r) { Log "  PostForm: Antwort war LEER (null)" }
        else { Log "  PostForm: Antwort erhalten" }
        return $r
    } catch {
        Log ("  PostForm EXCEPTION: " + $_.Exception.Message)
        # In PowerShell 5.1 steht der echte Fehler-Body in ErrorDetails.Message
        $bodyTxt = $null
        if ($_.ErrorDetails -and $_.ErrorDetails.Message) { $bodyTxt = $_.ErrorDetails.Message }
        elseif ($_.Exception.Response) {
            try { $sr = New-Object IO.StreamReader($_.Exception.Response.GetResponseStream()); $bodyTxt = $sr.ReadToEnd() } catch {}
        }
        Log ("  PostForm Antwort-Body: " + $bodyTxt)
        if ($bodyTxt) {
            try { return ($bodyTxt | ConvertFrom-Json) }
            catch { return [pscustomobject]@{ error = "http_error"; error_description = $bodyTxt } }
        }
        return [pscustomobject]@{ error = "net_error"; error_description = $_.Exception.Message }
    }
}

function GraphGet($url, $token) {
    return Invoke-RestMethod -Uri $url -Headers @{ Authorization = "Bearer $token" } -TimeoutSec 30
}

function Decode-Jwt($token) {
    # Liest Name/E-Mail direkt aus dem Zugangs-Token (keine extra Berechtigung noetig).
    try {
        $payload = $token.Split('.')[1].Replace('-', '+').Replace('_', '/')
        switch ($payload.Length % 4) { 2 { $payload += '==' } 3 { $payload += '=' } }
        $json = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($payload))
        return ($json | ConvertFrom-Json)
    } catch { return $null }
}

function GraphSendMail($token, $payload) {
    try {
        $json = ($payload | ConvertTo-Json -Depth 20)
        Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/me/sendMail" `
            -Headers @{ Authorization = "Bearer $token" } `
            -ContentType "application/json; charset=utf-8" `
            -Body ([Text.Encoding]::UTF8.GetBytes($json)) -TimeoutSec 60 | Out-Null
        return @{ ok = $true }
    } catch {
        $msg = $_.Exception.Message
        try {
            $resp = $_.Exception.Response
            if ($resp) { $sr = New-Object IO.StreamReader($resp.GetResponseStream()); $msg = $sr.ReadToEnd() }
        } catch {}
        return @{ ok = $false; error = $msg }
    }
}

function Normalize-DeferUntil($value) {
    # Gibt einen UTC-Zeitstempel im Graph-SystemTime-Format zurück oder $null.
    # Wichtig: Bei ungültigem/vergangenem Zeitpunkt wird NICHT sofort gesendet.
    if ($null -eq $value) { return $null }
    $raw = ([string]$value).Trim()
    if (-not $raw) { return $null }
    try {
        $dto = [DateTimeOffset]::Parse($raw, [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind)
        if ($dto.ToUniversalTime() -le [DateTimeOffset]::UtcNow.AddSeconds(30)) {
            throw "Le moment du rappel est déjà passé."
        }
        return $dto.UtcDateTime.ToString("yyyy-MM-ddTHH:mm:ssZ", [Globalization.CultureInfo]::InvariantCulture)
    } catch {
        throw ("Moment de rappel invalide: " + $raw)
    }
}

# ----------------------------------------------------------- Token
function Load-Tokens {
    if (Test-Path $TokenFile) {
        try { return (Get-Content -Raw -Path $TokenFile | ConvertFrom-Json) } catch { return $null }
    }
    return $null
}
function Save-Tokens($tok, $account) {
    $data = @{
        access_token  = $tok.access_token
        refresh_token = $tok.refresh_token
        expires_at    = ([DateTimeOffset]::UtcNow.ToUnixTimeSeconds() + [int]$tok.expires_in - 60)
    }
    if ($account) { $data.account = $account }
    else { $old = Load-Tokens; if ($old -and $old.account) { $data.account = $old.account } }
    ($data | ConvertTo-Json -Depth 6) | Set-Content -Path $TokenFile -Encoding UTF8
}
function Get-AccessToken {
    $t = Load-Tokens
    if (-not $t) { return $null }
    $now = [DateTimeOffset]::UtcNow.ToUnixTimeSeconds()
    if ($t.access_token -and $now -lt [int64]$t.expires_at) { return $t.access_token }
    if (-not $t.refresh_token) { return $null }
    $tok = PostForm "$Base/token" @{ grant_type = "refresh_token"; client_id = $ClientId; refresh_token = $t.refresh_token; scope = $Scope }
    if ($tok.access_token) { Save-Tokens $tok $null; return $tok.access_token }
    return $null
}

# ----------------------------------------------------------- Signatur (Windows)
function HtmlToText($html) {
    $s = $html
    $s = [regex]::Replace($s, '(?is)<!--.*?-->', '')
    $s = [regex]::Replace($s, '(?is)<head.*?</head>', '')
    $s = [regex]::Replace($s, '(?is)<style.*?</style>', '')
    $s = [regex]::Replace($s, '(?i)<br\s*/?>', "`n")
    $s = [regex]::Replace($s, '(?i)</(p|div|tr|li|h[1-6])\s*>', "`n")
    $s = [regex]::Replace($s, '(?s)<[^>]+>', '')
    $s = [System.Net.WebUtility]::HtmlDecode($s)
    $lines = @()
    foreach ($ln in ($s -split "`n")) {
        $t = ($ln -replace '[ \t\xA0]+', ' ').Trim()
        if ($t -eq '' -and ($lines.Count -eq 0 -or $lines[-1] -eq '')) { continue }
        $lines += $t
    }
    return (($lines -join "`n").Trim())
}
function Get-Signatures {
    $dir = Join-Path $env:APPDATA "Microsoft\Signatures"
    $out = @()
    if (Test-Path $dir) {
        Get-ChildItem -Path $dir -Filter *.htm -File | ForEach-Object {
            try {
                $html = [System.IO.File]::ReadAllText($_.FullName)
                $out += @{ name = $_.BaseName; text = (HtmlToText $html) }
            } catch {}
        }
    }
    return $out
}
function Find-Logo {
    $dir = Join-Path $env:APPDATA "Microsoft\Signatures"
    Log ("Find-Logo: Ordner = " + $dir + " | existiert: " + (Test-Path $dir))
    if (-not (Test-Path $dir)) { return $null }
    # Robust: alle Dateien rekursiv, dann nach Bild-Endung filtern (kein -Include-Quirk)
    $imgs = Get-ChildItem -Path $dir -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.Extension -match '(?i)\.(png|jpe?g|gif|bmp)$' }
    Log ("Find-Logo: " + @($imgs).Count + " Bilddatei(en) gefunden")
    $best = $null; $bestSize = 0
    foreach ($f in @($imgs)) {
        Log ("   - " + $f.FullName + "  (" + $f.Length + " Bytes)")
        if ($f.Length -gt 1000 -and $f.Length -gt $bestSize) { $best = $f; $bestSize = $f.Length }
    }
    if (-not $best) { Log "Find-Logo: kein geeignetes Bild (>1000 Bytes)"; return $null }
    Log ("Find-Logo: gewaehlt -> " + $best.Name + " (" + $best.Length + " Bytes)")
    $bytes = [IO.File]::ReadAllBytes($best.FullName)
    $ext = $best.Extension.TrimStart('.').ToLower()
    $mime = if ($ext -eq 'jpg' -or $ext -eq 'jpeg') { 'image/jpeg' } else { "image/$ext" }
    $b64 = [Convert]::ToBase64String($bytes)
    return @{ name = $best.Name; mime = $mime; dataUrl = "data:$mime;base64,$b64" }
}

# ----------------------------------------------------------- Signatur aus gesendeter Mail (Graph)
function Graph-LogoFromMail($token) {
    try {
        $dollar = [char]36
        $u = "https://graph.microsoft.com/v1.0/me/mailFolders/SentItems/messages?" + $dollar + "top=30&" + $dollar + "select=id&" + $dollar + "filter=hasAttachments%20eq%20true"
        $msgs = (GraphGet $u $token).value
        Log ("Graph-Logo: " + @($msgs).Count + " gesendete Mails mit Anhang")
        $countByKey = @{}; $store = @{}
        foreach ($m in $msgs) {
            $au = "https://graph.microsoft.com/v1.0/me/messages/" + $m.id + "/attachments?" + $dollar + "select=id,name,contentType,isInline,size"
            $atts = (GraphGet $au $token).value
            foreach ($a in $atts) {
                if (($a.contentType -like "image/*") -and ($a.size -gt 1500)) {
                    $key = [string]$a.size
                    if (-not $countByKey.ContainsKey($key)) { $countByKey[$key] = 0; $store[$key] = @{ mid = $m.id; aid = $a.id; name = $a.name; mime = $a.contentType } }
                    $countByKey[$key] = $countByKey[$key] + 1
                }
            }
        }
        $bestKey = $null; $bestCount = 0
        foreach ($k in $countByKey.Keys) { if ($countByKey[$k] -gt $bestCount) { $bestCount = $countByKey[$k]; $bestKey = $k } }
        if (-not $bestKey) { Log "Graph-Logo: kein Inline-Bild gefunden"; return $null }
        $sel = $store[$bestKey]
        $full = GraphGet ("https://graph.microsoft.com/v1.0/me/messages/" + $sel.mid + "/attachments/" + $sel.aid) $token
        if (-not $full.contentBytes) { return $null }
        Log ("Graph-Logo: gewaehlt " + $sel.name + " (" + $bestKey + " Bytes, " + $bestCount + "x)")
        return @{ name = $sel.name; mime = $sel.mime; dataUrl = ("data:" + $sel.mime + ";base64," + $full.contentBytes) }
    } catch { Log ("Graph-Logo FEHLER: " + $_.Exception.Message); return $null }
}

function Graph-SignatureFromMail($token) {
    try {
        $dollar = [char]36
        $u = "https://graph.microsoft.com/v1.0/me/mailFolders/SentItems/messages?" + $dollar + "top=15&" + $dollar + "select=id,body"
        $msgs = (GraphGet $u $token).value
        $keywords = @("Bien cordialement", "Cordialement", "Mit freundlichen", "Freundliche", "Beste Gruesse", "Liebe Gruesse", "Bien a vous", "Sportlechen")
        foreach ($m in $msgs) {
            $html = $m.body.content
            if (-not $html) { continue }
            $txt = HtmlToText $html
            $idx = -1
            foreach ($kw in $keywords) { $p = $txt.LastIndexOf($kw); if ($p -gt $idx) { $idx = $p } }
            if ($idx -ge 0) {
                $sig = $txt.Substring($idx)
                $lines = ($sig -split "`n") | Select-Object -First 12
                $res = (($lines -join "`n").Trim())
                if ($res.Length -gt 3) { Log "Graph-Signatur: Text aus gesendeter Mail extrahiert"; return $res }
            }
        }
        Log "Graph-Signatur: kein Signaturblock gefunden"
        return $null
    } catch { Log ("Graph-Signatur FEHLER: " + $_.Exception.Message); return $null }
}

# ----------------------------------------------------------- HTTP-Server
function Find-HeaderEnd($arr, $len) {
    for ($i = 0; $i -lt $len - 3; $i++) {
        if ($arr[$i] -eq 13 -and $arr[$i+1] -eq 10 -and $arr[$i+2] -eq 13 -and $arr[$i+3] -eq 10) { return $i }
    }
    return -1
}
function Read-Request($stream) {
    try { $stream.ReadTimeout = 15000 } catch {}
    $buf = New-Object byte[] 16384
    $ms = New-Object System.IO.MemoryStream
    $headerEnd = -1
    while ($true) {
        $arr = $ms.ToArray()
        $headerEnd = Find-HeaderEnd $arr $arr.Length
        if ($headerEnd -ge 0) { break }
        $n = $stream.Read($buf, 0, $buf.Length)
        if ($n -le 0) { break }
        $ms.Write($buf, 0, $n)
        if ($ms.Length -gt 5000000) { break }
    }
    $arr = $ms.ToArray()
    if ($headerEnd -lt 0) { return $null }
    $headerText = [Text.Encoding]::ASCII.GetString($arr, 0, $headerEnd)
    $contentLength = 0
    foreach ($line in ($headerText -split "`r`n")) {
        if ($line -match '^(?i)Content-Length:\s*(\d+)') { $contentLength = [int]$matches[1] }
    }
    $bodyStart = $headerEnd + 4
    while (($arr.Length - $bodyStart) -lt $contentLength) {
        $n = $stream.Read($buf, 0, $buf.Length)
        if ($n -le 0) { break }
        $ms.Write($buf, 0, $n)
        $arr = $ms.ToArray()
        if ($ms.Length -gt 50000000) { break }
    }
    $arr = $ms.ToArray()
    $body = ""
    if ($contentLength -gt 0 -and $arr.Length -ge $bodyStart) {
        $blen = [Math]::Min($contentLength, $arr.Length - $bodyStart)
        $body = [Text.Encoding]::UTF8.GetString($arr, $bodyStart, $blen)
    }
    $firstLine = ($headerText -split "`r`n")[0]
    $parts = $firstLine -split ' '
    return [pscustomobject]@{ Method = $parts[0]; Path = $parts[1]; Body = $body }
}
function Send-Bytes($stream, $status, $contentType, [byte[]]$body) {
    $head = "HTTP/1.1 $status`r`nContent-Type: $contentType`r`nContent-Length: $($body.Length)`r`nConnection: close`r`n`r`n"
    $hb = [Text.Encoding]::ASCII.GetBytes($head)
    $stream.Write($hb, 0, $hb.Length)
    if ($body.Length -gt 0) { $stream.Write($body, 0, $body.Length) }
    $stream.Flush()
}
function Send-Json($stream, $obj) {
    $json = ($obj | ConvertTo-Json -Depth 12 -Compress)
    Send-Bytes $stream "200 OK" "application/json; charset=utf-8" ([Text.Encoding]::UTF8.GetBytes($json))
}

function Handle-Request($stream, $req) {
    $path = ($req.Path -split '\?')[0]
    $method = $req.Method
    Write-Host ("[" + $method + "] " + $path)
    try {

    # ---- API ----
    if ($path -eq "/api/graph/capabilities") {
        Send-Json $stream @{ ok = $true; deferredSend = $true; platform = "windows-powershell"; appVersion = 81 }
        return
    }
    if ($path -eq "/api/graph/account") {
        $t = Load-Tokens
        if ($t -and $t.account -and (Get-AccessToken)) { Send-Json $stream @{ ok = $true; signedIn = $true; account = $t.account } }
        else { Send-Json $stream @{ ok = $true; signedIn = $false } }
        return
    }
    if ($path -eq "/api/graph/login-start") {
        $dc = PostForm "$Base/devicecode" @{ client_id = $ClientId; scope = $Scope }
        if ($dc.device_code) {
            Write-Host ("  login-start: Code erhalten (" + $dc.user_code + ")")
            $script:Pending = $dc
            Send-Json $stream @{ ok = $true; user_code = $dc.user_code; verification_uri = $dc.verification_uri; message = $dc.message }
        } else {
            $emsg = (("" + $dc.error_description + " " + $dc.error)).Trim()
            if (-not $emsg) { $emsg = "Unerwartete Antwort: " + ($dc | ConvertTo-Json -Compress) }
            Write-Host ("  login-start FEHLER: " + $emsg)
            Send-Json $stream @{ ok = $false; error = $emsg }
        }
        return
    }
    if ($path -eq "/api/graph/login-poll") {
        if (-not $script:Pending) { Send-Json $stream @{ ok = $false; error = "Kein Login gestartet." }; return }
        $tok = PostForm "$Base/token" @{ grant_type = "urn:ietf:params:oauth:grant-type:device_code"; client_id = $ClientId; device_code = $script:Pending.device_code }
        if ($tok.access_token) {
            $acct = @{ name = ""; email = "" }
            try {
                $c = Decode-Jwt $tok.access_token
                if ($c) {
                    $acct.name = $c.name
                    $acct.email = if ($c.upn) { $c.upn } elseif ($c.preferred_username) { $c.preferred_username } else { $c.unique_name }
                }
            } catch {}
            Save-Tokens $tok $acct
            $script:Pending = $null
            Log ("login OK: " + $acct.email)
            Send-Json $stream @{ ok = $true; status = "ok"; account = $acct }
        } elseif ($tok.error -eq "authorization_pending" -or $tok.error -eq "slow_down") {
            Send-Json $stream @{ ok = $true; status = "pending" }
        } else {
            $script:Pending = $null
            Send-Json $stream @{ ok = $true; status = "error"; error = ("" + $tok.error_description + $tok.error) }
        }
        return
    }
    if ($path -eq "/api/graph/logout") {
        try { if (Test-Path $TokenFile) { Remove-Item $TokenFile -Force } } catch {}
        $script:Pending = $null
        Send-Json $stream @{ ok = $true }
        return
    }
    if ($path -eq "/api/graph/send") {
        $token = Get-AccessToken
        if (-not $token) { Send-Json $stream @{ ok = $false; error = "Nicht angemeldet. Bitte erneut mit Microsoft anmelden." }; return }
        $data = $null
        try { $data = $req.Body | ConvertFrom-Json } catch {}
        if (-not $data -or -not $data.messages) { Send-Json $stream @{ ok = $false; error = "Aucun e-mail." }; return }
        $results = @()
        foreach ($m in $data.messages) {
            $msg = @{
                subject      = $m.subject
                body         = @{ contentType = "HTML"; content = $m.html }
                toRecipients = @($m.to | ForEach-Object { @{ emailAddress = @{ address = $_ } } })
            }
            $deferUntil = $null
            try {
                if ($m.PSObject.Properties.Name -contains "deferUntil" -and $m.deferUntil) {
                    $deferUntil = Normalize-DeferUntil $m.deferUntil
                    $msg["singleValueExtendedProperties"] = @(@{
                        id    = "SystemTime 0x3FEF"
                        value = $deferUntil
                    })
                    Log ("Deferred send requested for message " + $m.id + " -> " + $deferUntil)
                }
            } catch {
                $results += @{ id = $m.id; ok = $false; error = $_.Exception.Message }
                continue
            }
            if ($data.logo -and $data.logo.contentBytes -and ($m.html -match 'cid:siglogo')) {
                $msg["attachments"] = @(@{
                    "@odata.type" = "#microsoft.graph.fileAttachment"
                    name          = $data.logo.name
                    contentType   = $data.logo.mime
                    contentBytes  = $data.logo.contentBytes
                    isInline      = $true
                    contentId     = "siglogo"
                })
            }
            $r = GraphSendMail $token @{ message = $msg; saveToSentItems = $true }
            if ($r.ok) { $results += @{ id = $m.id; ok = $true; deferredUntil = $deferUntil } }
            else { $results += @{ id = $m.id; ok = $false; error = $r.error } }
        }
        Send-Json $stream @{ ok = $true; results = $results }
        return
    }
    if ($path -eq "/api/outlook-signatures") {
        try {
            $sigs = Get-Signatures
            if (-not $sigs -or @($sigs).Count -eq 0) {
                $token = Get-AccessToken
                if ($token) {
                    $t = Graph-SignatureFromMail $token
                    if ($t) { $sigs = @(@{ name = "Depuis un e-mail envoyé"; text = $t }) }
                }
            }
            Send-Json $stream @{ ok = $true; signatures = $sigs }
        } catch { Send-Json $stream @{ ok = $false; error = $_.Exception.Message } }
        return
    }
    if ($path -eq "/api/find-logo") {
        try {
            $logo = Find-Logo
            if (-not $logo) {
                $token = Get-AccessToken
                if ($token) { $logo = Graph-LogoFromMail $token }
            }
            if ($logo) { Send-Json $stream @{ ok = $true; logo = $logo } }
            else { Send-Json $stream @{ ok = $false; error = "Kein Logo gefunden." } }
        } catch { Send-Json $stream @{ ok = $false; error = $_.Exception.Message } }
        return
    }

    # ---- statische Dateien ----
    if ($path -eq "/" -or $path -eq "") { $path = "/graph.html" }
    $rel = $path.TrimStart('/')
    $file = Join-Path $ScriptDir $rel
    if ((Test-Path $file -PathType Leaf)) {
        $ext = [IO.Path]::GetExtension($file).ToLower()
        $ct = switch ($ext) {
            ".html" { "text/html; charset=utf-8" }
            ".js"   { "application/javascript; charset=utf-8" }
            ".css"  { "text/css; charset=utf-8" }
            ".png"  { "image/png" }
            ".jpg"  { "image/jpeg" }
            ".svg"  { "image/svg+xml" }
            default { "application/octet-stream" }
        }
        Send-Bytes $stream "200 OK" $ct ([IO.File]::ReadAllBytes($file))
    } else {
        Send-Bytes $stream "404 Not Found" "text/plain; charset=utf-8" ([Text.Encoding]::UTF8.GetBytes("Not found"))
    }

    } catch {
        Write-Host ("  SERVERFEHLER: " + $_.Exception.Message)
        try { Send-Json $stream @{ ok = $false; error = ("Serverfehler: " + $_.Exception.Message) } } catch {}
    }
}

# ----------------------------------------------------------- Start
# Interface française locale : pas de mise à jour GitHub automatique.
try {
    $listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $Port)
    $listener.Start()
} catch {
    Write-Host ""
    Write-Host "ERREUR : impossible de démarrer le serveur local (port $Port déjà utilisé ?)."
    Write-Host $_.Exception.Message
    Read-Host "Entrée pour fermer"
    return
}

$url = "http://127.0.0.1:$Port/graph.html"
try { Set-Content -Path $LogFile -Value ("=== Start " + (Get-Date) + " ===") -Encoding UTF8 } catch {}
Write-Host "============================================================"
Write-Host "  Rendez-vous parents est lancé.   [Version : v81 - FR]"
Write-Host "  Dans le navigateur :  $url"
Write-Host "  Laissez cette fenêtre ouverte. La fermer = quitter."
Write-Host "============================================================"
try { Start-Process $url } catch {}

while ($true) {
    $client = $null
    try {
        $client = $listener.AcceptTcpClient()
        $stream = $client.GetStream()
        $req = Read-Request $stream
        if ($req) { Handle-Request $stream $req }
    } catch {
        # einzelne fehlerhafte Anfrage ignorieren, Server weiterlaufen lassen
    } finally {
        if ($client) { try { $client.Close() } catch {} }
    }
}
