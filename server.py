#!/usr/bin/env python3
"""
Lokaler Helfer fuer die ElternTermine-Browser-App.

- Liefert die index.html aus
- Nimmt unter POST /api/send die fertigen E-Mails entgegen und
  verschickt sie ueber Gmail (SMTP, SSL) als BCC.

Keine Zusatz-Pakete noetig (nur Python-Standardbibliothek).
Start:  python3 server.py
"""

import http.server
import socketserver
import json
import smtplib
import ssl
import sys
import webbrowser
import threading
import os
import time
import subprocess
import urllib.request
import urllib.parse
import urllib.error
import errno
from email.message import EmailMessage

PORT = 8765


def _port_is_in_use_error(exc):
    return isinstance(exc, OSError) and getattr(exc, "errno", None) in (errno.EADDRINUSE, 48, 98, 10048)


def _make_server_with_fallback(host, preferred_port, handler_cls):
    """Startet auf preferred_port; falls belegt, nimmt automatisch einen freien Port."""
    try:
        return Server((host, preferred_port), handler_cls), preferred_port, False
    except OSError as exc:
        if not _port_is_in_use_error(exc):
            raise
        # Port 0 laesst macOS/Windows/Linux automatisch einen freien Port waehlen.
        httpd = Server((host, 0), handler_cls)
        return httpd, httpd.server_address[1], True

# Als .exe (PyInstaller) gebuendelt liegen die HTML-Dateien im temporaeren
# Entpackordner (_MEIPASS); der Login-Schluessel muss aber an einen dauerhaften,
# beschreibbaren Ort (neben die .exe). Als normales Skript: alles im Skriptordner.
if getattr(sys, "frozen", False):
    RES_DIR = sys._MEIPASS                          # ausgelieferte Dateien (read-only)
    DATA_DIR = os.path.dirname(sys.executable)      # beschreibbar, neben der .exe
else:
    RES_DIR = os.path.dirname(os.path.abspath(__file__))
    DATA_DIR = RES_DIR
DIRECTORY = RES_DIR


def _as_line(s):
    """Einzeiliger String (z. B. Betreff) fuer ein AppleScript-Literal."""
    return (s or "").replace("\\", "\\\\").replace('"', '\\"').replace("\n", " ").replace("\r", " ")


def _as_text(s):
    """Mehrzeiliger Text fuer ein AppleScript-Literal; Zeilenumbrueche bleiben erhalten."""
    s = (s or "").replace("\\", "\\\\").replace('"', '\\"')
    s = s.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\\n")
    return s


def build_outlook_script(subject, body, to_list, bcc_list, action):
    """AppleScript-Zeilen, die in Microsoft Outlook (Mac) eine Mail erstellen
    und je nach action senden ('send') oder zur Kontrolle oeffnen ('open').
    Inhalt wird als REINER TEXT gesetzt (Outlook fuer Mac rendert hier kein HTML)."""
    lines = ['tell application "Microsoft Outlook"']
    lines.append(
        'set newMsg to make new outgoing message with properties '
        '{{subject:"{s}", content:"{c}"}}'.format(s=_as_line(subject), c=_as_text(body))
    )
    for a in (to_list or []):
        lines.append(
            'make new to recipient at newMsg with properties '
            '{{email address:{{address:"{a}"}}}}'.format(a=_as_line(a))
        )
    for b in (bcc_list or []):
        lines.append(
            'make new bcc recipient at newMsg with properties '
            '{{email address:{{address:"{a}"}}}}'.format(a=_as_line(b))
        )
    lines.append("send newMsg" if action == "send" else "open newMsg")
    lines.append("end tell")
    return lines


def html_to_text(s):
    """Word/Outlook-HTML-Signatur in sauberen Klartext umwandeln."""
    import re
    import html as _html
    s = re.sub(r"(?is)<!--.*?-->", "", s)
    s = re.sub(r"(?is)<head.*?</head>", "", s)
    s = re.sub(r"(?is)<style.*?</style>", "", s)
    s = re.sub(r"(?i)<br\s*/?>", "\n", s)
    s = re.sub(r"(?i)</(p|div|tr|li|h[1-6])\s*>", "\n", s)
    s = re.sub(r"(?s)<[^>]+>", "", s)
    s = _html.unescape(s).replace("\xa0", " ")
    out = []
    for ln in s.split("\n"):
        ln = re.sub(r"[ \t]+", " ", ln).strip()
        if ln == "" and (not out or out[-1] == ""):
            continue
        out.append(ln)
    while out and out[-1] == "":
        out.pop()
    while out and out[0] == "":
        out.pop(0)
    return "\n".join(out)


def _signatures_mac():
    names_script = (
        'set out to ""\n'
        'tell application "Microsoft Outlook"\n'
        '  repeat with s in signatures\n'
        '    set out to out & (name of s) & linefeed\n'
        '  end repeat\n'
        'end tell\n'
        'return out'
    )
    proc = subprocess.run(["osascript", "-e", names_script],
                          capture_output=True, text=True, timeout=30)
    if proc.returncode != 0:
        raise RuntimeError(proc.stderr.strip() or "Signatures illisibles")
    names = [n.strip() for n in proc.stdout.splitlines() if n.strip()]
    sigs = []
    for name in names:
        cp = subprocess.run(
            ["osascript", "-e",
             'tell application "Microsoft Outlook" to get content of signature "{n}"'.format(
                 n=name.replace("\\", "\\\\").replace('"', '\\"'))],
            capture_output=True, text=True, timeout=30)
        text = html_to_text(cp.stdout) if cp.returncode == 0 else ""
        sigs.append({"name": name, "text": text})
    return sigs


def _signatures_windows():
    """Klassisches Outlook fuer Windows: Signatur-.htm-Dateien lesen."""
    base = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Signatures")
    if not os.path.isdir(base):
        return []
    sigs = []
    for fn in sorted(os.listdir(base)):
        if not fn.lower().endswith(".htm") and not fn.lower().endswith(".html"):
            continue
        path = os.path.join(base, fn)
        raw = b""
        try:
            with open(path, "rb") as f:
                raw = f.read()
        except Exception:
            continue
        html = None
        for enc in ("utf-8", "cp1252", "utf-16", "latin-1"):
            try:
                html = raw.decode(enc)
                break
            except Exception:
                continue
        if html is None:
            html = raw.decode("utf-8", "replace")
        sigs.append({"name": os.path.splitext(fn)[0], "text": html_to_text(html)})
    return sigs


def get_outlook_signatures():
    """Liste der Outlook-Signaturen als [{name, text}] (Inhalt als Klartext).
    Funktioniert auf macOS (AppleScript) und Windows (Signatur-Dateien)."""
    if os.name == "nt":
        return _signatures_windows()
    return _signatures_mac()


def find_signature_logo():
    """Versucht, das Logo-Bild der Signatur automatisch zu finden.
    Windows: groesstes Bild im Outlook-Signaturordner.
    macOS:  haeufigstes kleines Bild in den Apple-Mail-Anhaengen (Signatur-Logo
            wiederholt sich in vielen Mails)."""
    import base64
    import glob
    import hashlib
    import collections

    def as_logo(path, data, count=None):
        ext = path.lower().rsplit(".", 1)[-1] if "." in path else "png"
        mime = "image/" + ("jpeg" if ext in ("jpg", "jpeg") else ext)
        return {"name": "logo." + ext, "mime": mime,
                "dataUrl": "data:%s;base64,%s" % (mime, base64.b64encode(data).decode()),
                "count": count}

    if os.name == "nt":
        base = os.path.join(os.environ.get("APPDATA", ""), "Microsoft", "Signatures")
        best, best_sz = None, 0
        for root, _dirs, files in os.walk(base):
            for fn in files:
                ext = fn.lower().rsplit(".", 1)[-1] if "." in fn else ""
                if ext in ("png", "jpg", "jpeg", "gif"):
                    p = os.path.join(root, fn)
                    try:
                        sz = os.path.getsize(p)
                    except Exception:
                        continue
                    if 1500 < sz and sz > best_sz:
                        best, best_sz = p, sz
        if best:
            with open(best, "rb") as f:
                return as_logo(best, f.read())
        return None

    # macOS: haeufigstes kleines Bild in Apple-Mail
    files = []
    for ext in ("png", "jpg", "jpeg"):
        files += glob.glob(os.path.expanduser("~/Library/Mail/**/image00*." + ext), recursive=True)
    counter = collections.Counter()
    info = {}
    for f in files[:4000]:
        try:
            sz = os.path.getsize(f)
            if sz < 3000 or sz > 300000:
                continue
            data = open(f, "rb").read()
            hh = hashlib.md5(data).hexdigest()
            counter[hh] += 1
            info[hh] = (f, data)
        except Exception:
            continue
    if not counter:
        return None
    hh, n = counter.most_common(1)[0]
    if n < 2:
        return None
    f, data = info[hh]
    return as_logo(f, data, n)


## ===================== Microsoft Graph (Geraetecode-Login) =====================
# Oeffentlicher Microsoft-Client "Microsoft Graph Command Line Tools" -> keine
# eigene App-Registrierung / keine IT noetig. Jeder meldet sich einmal an.
GRAPH_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
GRAPH_TENANT = "organizations"
GRAPH_SCOPE = ("https://graph.microsoft.com/Mail.Send "
               "https://graph.microsoft.com/User.Read offline_access")
OAUTH_BASE = "https://login.microsoftonline.com/" + GRAPH_TENANT + "/oauth2/v2.0"
TOKEN_FILE = os.path.join(DATA_DIR, "graph_token.json")

_pending_device = {}   # zwischen Login-Start und Polling


def _gctx():
    return make_ssl_context()


def _form_post(url, data):
    body = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(url, body, {"Content-Type": "application/x-www-form-urlencoded"})
    try:
        with urllib.request.urlopen(req, context=_gctx(), timeout=30) as r:
            return r.status, json.load(r)
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.load(e)
        except Exception:
            return e.code, {"error": "http_" + str(e.code)}


def _graph_get(url, token):
    req = urllib.request.Request(url, headers={"Authorization": "Bearer " + token})
    with urllib.request.urlopen(req, context=_gctx(), timeout=30) as r:
        return json.load(r)


def _graph_post(url, token, payload):
    body = json.dumps(payload).encode()
    req = urllib.request.Request(url, body, {
        "Authorization": "Bearer " + token, "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, context=_gctx(), timeout=30) as r:
            return r.status, ""
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "replace")


def _save_tokens(tok, account=None):
    data = {
        "access_token": tok.get("access_token"),
        "refresh_token": tok.get("refresh_token"),
        "expires_at": time.time() + int(tok.get("expires_in", 3600)) - 60,
    }
    if account:
        data["account"] = account
    else:
        old = _load_tokens()
        if old and old.get("account"):
            data["account"] = old["account"]
    try:
        with open(TOKEN_FILE, "w") as f:
            json.dump(data, f)
    except Exception:
        pass


def _load_tokens():
    try:
        with open(TOKEN_FILE) as f:
            return json.load(f)
    except Exception:
        return None


def _access_token():
    """Gibt ein gueltiges Access-Token zurueck (erneuert bei Bedarf) oder None."""
    t = _load_tokens()
    if not t:
        return None
    if t.get("access_token") and time.time() < t.get("expires_at", 0):
        return t["access_token"]
    rt = t.get("refresh_token")
    if not rt:
        return None
    st, tok = _form_post(OAUTH_BASE + "/token", {
        "grant_type": "refresh_token", "client_id": GRAPH_CLIENT_ID,
        "refresh_token": rt, "scope": GRAPH_SCOPE})
    if "access_token" in tok:
        _save_tokens(tok)
        return tok["access_token"]
    return None


def make_ssl_context():
    """SSL-Kontext mit zuverlaessigem Zertifikatsspeicher.

    Auf macOS fehlt bei frischen python.org-Installationen oft die
    cert.pem-Datei (das beruechtigte 'Install Certificates'-Problem).
    certifi ist bei diesen Installationen aber mitgeliefert, daher
    nutzen wir dessen CA-Buendel, sonst den Standard.
    """
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=DIRECTORY, **kwargs)

    def log_message(self, fmt, *args):
        # Etwas ruhiger im Terminal
        pass

    def do_GET(self):
        if self.path in ("/", "/index.html"):
            self.send_response(302)
            self.send_header("Location", "/graph.html")
            self.end_headers()
            return
        if self.path == "/api/outlook-signatures":
            return self.handle_signatures()
        if self.path == "/api/graph/account":
            return self.handle_graph_account()
        if self.path == "/api/find-logo":
            return self.handle_find_logo()
        return super().do_GET()

    def handle_find_logo(self):
        try:
            logo = find_signature_logo()
            if logo:
                return self._json(200, {"ok": True, "logo": logo})
            return self._json(200, {"ok": False, "error": "Aucun logo trouvé."})
        except Exception as e:
            return self._json(200, {"ok": False, "error": str(e)})

    def _read_json(self):
        length = int(self.headers.get("Content-Length", 0))
        try:
            return json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception:
            return {}

    # ---- Graph: Konto-Status ----
    def handle_graph_account(self):
        t = _load_tokens()
        if t and t.get("account") and _access_token():
            return self._json(200, {"ok": True, "signedIn": True, "account": t["account"]})
        return self._json(200, {"ok": True, "signedIn": False})

    # ---- Graph: Login starten (Geraetecode) ----
    def handle_graph_login_start(self):
        st, dc = _form_post(OAUTH_BASE + "/devicecode",
                            {"client_id": GRAPH_CLIENT_ID, "scope": GRAPH_SCOPE})
        if "device_code" not in dc:
            return self._json(200, {"ok": False, "error": dc.get("error_description") or dc.get("error") or "Erreur"})
        _pending_device["device_code"] = dc["device_code"]
        _pending_device["interval"] = int(dc.get("interval", 5))
        return self._json(200, {"ok": True,
                                "user_code": dc.get("user_code"),
                                "verification_uri": dc.get("verification_uri"),
                                "message": dc.get("message")})

    # ---- Graph: Login abfragen (Polling) ----
    def handle_graph_login_poll(self):
        dcode = _pending_device.get("device_code")
        if not dcode:
            return self._json(200, {"ok": False, "error": "Aucune connexion démarrée."})
        st, tok = _form_post(OAUTH_BASE + "/token", {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": GRAPH_CLIENT_ID, "device_code": dcode})
        if "access_token" in tok:
            _save_tokens(tok)
            account = {"name": "", "email": ""}
            try:
                me = _graph_get("https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName",
                                tok["access_token"])
                account = {"name": me.get("displayName", ""),
                           "email": me.get("mail") or me.get("userPrincipalName", "")}
            except Exception:
                pass
            _save_tokens(tok, account)
            _pending_device.clear()
            return self._json(200, {"ok": True, "status": "ok", "account": account})
        err = tok.get("error")
        if err in ("authorization_pending", "slow_down"):
            return self._json(200, {"ok": True, "status": "pending"})
        _pending_device.clear()
        return self._json(200, {"ok": True, "status": "error",
                                "error": tok.get("error_description") or err or "Erreur"})

    # ---- Graph: Abmelden ----
    def handle_graph_logout(self):
        try:
            if os.path.exists(TOKEN_FILE):
                os.remove(TOKEN_FILE)
        except Exception:
            pass
        _pending_device.clear()
        return self._json(200, {"ok": True})

    # ---- Graph: Senden ----
    def handle_graph_send(self):
        data = self._read_json()
        messages = data.get("messages", []) or []
        logo = data.get("logo")  # {name, mime, contentBytes} oder None
        if not messages:
            return self._json(200, {"ok": False, "error": "Aucun e-mail."})
        token = _access_token()
        if not token:
            return self._json(200, {"ok": False, "error": "Non connecté. Veuillez vous reconnecter avec Microsoft."})
        results = []
        for m in messages:
            payload = {
                "message": {
                    "subject": m.get("subject", ""),
                    "body": {"contentType": "HTML", "content": m.get("html", "")},
                    "toRecipients": [{"emailAddress": {"address": a}} for a in (m.get("to") or []) if a],
                },
                "saveToSentItems": True,
            }
            if logo and logo.get("contentBytes") and "cid:siglogo" in m.get("html", ""):
                payload["message"]["attachments"] = [{
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": logo.get("name", "logo.png"),
                    "contentType": logo.get("mime", "image/png"),
                    "contentBytes": logo["contentBytes"],
                    "isInline": True, "contentId": "siglogo",
                }]
            code, body = _graph_post("https://graph.microsoft.com/v1.0/me/sendMail", token, payload)
            if code == 202:
                results.append({"id": m.get("id"), "ok": True})
            else:
                results.append({"id": m.get("id"), "ok": False, "error": "HTTP " + str(code) + " " + body[:200]})
        return self._json(200, {"ok": True, "results": results})

    def handle_signatures(self):
        try:
            sigs = get_outlook_signatures()
            return self._json(200, {"ok": True, "signatures": sigs})
        except Exception as e:
            return self._json(200, {"ok": False, "error": str(e)})

    def do_POST(self):
        if self.path == "/api/send":
            self.handle_send()
        elif self.path == "/api/outlook-send":
            self.handle_outlook_send()
        elif self.path == "/api/graph/login-start":
            self.handle_graph_login_start()
        elif self.path == "/api/graph/login-poll":
            self.handle_graph_login_poll()
        elif self.path == "/api/graph/logout":
            self.handle_graph_logout()
        elif self.path == "/api/graph/send":
            self.handle_graph_send()
        else:
            self.send_error(404, "Not found")

    # ---------------------------------------------------------- Outlook (Mac)
    def handle_outlook_send(self):
        length = int(self.headers.get("Content-Length", 0))
        try:
            data = json.loads(self.rfile.read(length).decode("utf-8"))
        except Exception:
            return self._json(400, {"ok": False, "error": "Requête non valable."})

        from_addr = (data.get("from") or "").strip()
        action = data.get("action", "send")  # "send" oder "open"
        messages = data.get("messages", []) or []
        if not messages:
            return self._json(200, {"ok": False, "error": "Aucun e-mail."})

        results = []
        for m in messages:
            # Empfaenger direkt ins An-Feld (Einzelversand, kein BCC noetig)
            recipients = [r for r in (m.get("to") or m.get("bcc") or []) if r]
            script_lines = build_outlook_script(
                subject=m.get("subject", ""),
                body=m.get("body", ""),
                to_list=recipients,
                bcc_list=[],
                action=action,
            )
            args = ["osascript"]
            for line in script_lines:
                args += ["-e", line]
            try:
                proc = subprocess.run(args, capture_output=True, text=True, timeout=60)
                if proc.returncode == 0:
                    results.append({"id": m.get("id"), "ok": True})
                else:
                    err = (proc.stderr or "").strip() or "Erreur AppleScript"
                    results.append({"id": m.get("id"), "ok": False, "error": err})
            except Exception as e:
                results.append({"id": m.get("id"), "ok": False, "error": str(e)})
        return self._json(200, {"ok": True, "results": results})

    # ------------------------------------------------------------------ send
    def handle_send(self):
        length = int(self.headers.get("Content-Length", 0))
        raw = self.rfile.read(length)
        try:
            data = json.loads(raw.decode("utf-8"))
        except Exception:
            return self._json(400, {"ok": False, "error": "Requête non valable."})

        sender = data.get("sender", {}) or {}
        host = sender.get("host", "smtp.gmail.com")
        port = int(sender.get("port", 465))
        user = (sender.get("user") or sender.get("email") or "").strip()
        password = sender.get("password", "")
        from_name = sender.get("name", "")
        from_addr = (sender.get("email") or user).strip()
        messages = data.get("messages", []) or []

        if not user or not password:
            return self._json(200, {"ok": False,
                                    "error": "Adresse e-mail ou mot de passe d’application manquant."})
        if not messages:
            return self._json(200, {"ok": False, "error": "Aucun e-mail à envoyer."})

        results = []
        try:
            context = make_ssl_context()
            with smtplib.SMTP_SSL(host, port, context=context, timeout=30) as server:
                server.login(user, password)
                for m in messages:
                    try:
                        msg = EmailMessage()
                        msg["Subject"] = m.get("subject", "")
                        msg["From"] = (f"{from_name} <{from_addr}>"
                                       if from_name else from_addr)
                        # Empfaenger direkt ins An-Feld (Einzelversand)
                        recipients = [r for r in (m.get("to") or m.get("bcc") or []) if r]
                        if recipients:
                            msg["To"] = ", ".join(recipients)
                        msg.set_content(m.get("body", ""))
                        html = m.get("html")
                        if html:
                            msg.add_alternative(html, subtype="html")
                        server.send_message(msg)
                        results.append({"id": m.get("id"), "ok": True})
                    except Exception as e:
                        results.append({"id": m.get("id"), "ok": False,
                                        "error": str(e)})
            return self._json(200, {"ok": True, "results": results})
        except smtplib.SMTPAuthenticationError:
            return self._json(200, {"ok": False,
                                    "error": ("Échec de la connexion à Gmail. "
                                              "Vérifiez l’adresse de l’expéditeur et utilisez "
                                              "un mot de passe d’application (pas votre mot de passe "
                                              "habituel).")})
        except Exception as e:
            return self._json(200, {"ok": False, "error": str(e)})

    # ------------------------------------------------------------------ util
    def _json(self, code, obj):
        body = json.dumps(obj, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)


class Server(socketserver.ThreadingTCPServer):
    allow_reuse_address = True
    daemon_threads = True


GITHUB_UI_URL = "https://raw.githubusercontent.com/joelschartz/elterntermine/main/graph.html"


def update_ui_from_github():
    """Conserve l’interface française locale de cette version."""
    return


def main():
    os.chdir(DIRECTORY)
    update_ui_from_github()
    host = "127.0.0.1"
    with _make_server_with_fallback(host, PORT, Handler)[0] as httpd:
        actual_port = httpd.server_address[1]
        used_fallback = actual_port != PORT
        url = f"http://{host}:{actual_port}/graph.html"
        print("=" * 56)
        print("  Rendez-vous parents est lancé.")
        if used_fallback:
            print(f"  Remarque : le port {PORT} était déjà utilisé.")
            print(f"  Le port {actual_port} a donc été utilisé automatiquement.")
        print(f"  Ouvrir dans le navigateur :  {url}")
        print("  Pour quitter : fermez cette fenêtre (ou Ctrl+C)")
        print("=" * 56)
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\nTerminé.")


if __name__ == "__main__":
    main()
