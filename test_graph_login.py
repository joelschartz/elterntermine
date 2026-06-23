#!/usr/bin/env python3
"""
MINI-TEST: Funktioniert die Geraetecode-Anmeldung (OAuth) ohne eigene
App-Registrierung in eurem Microsoft-Tenant?

Es wird KEIN Passwort in dieses Programm eingegeben. Du meldest dich auf
Microsofts echter Login-Seite an. Danach versucht das Skript, EINE Testmail
an dich selbst zu senden.

Start:  python3 test_graph_login.py
"""

import json
import time
import ssl
import urllib.request
import urllib.parse
import urllib.error

# Oeffentlicher Microsoft-Client "Microsoft Graph Command Line Tools"
# (von Microsoft bereitgestellt, erlaubt Geraetecode-Anmeldung).
CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"
TENANT = "organizations"
SCOPE = "https://graph.microsoft.com/Mail.Send https://graph.microsoft.com/User.Read offline_access"
BASE = "https://login.microsoftonline.com/" + TENANT + "/oauth2/v2.0"


def ssl_ctx():
    try:
        import certifi
        return ssl.create_default_context(cafile=certifi.where())
    except Exception:
        return ssl.create_default_context()


CTX = ssl_ctx()


def post_form(url, data):
    body = urllib.parse.urlencode(data).encode()
    req = urllib.request.Request(url, body, {"Content-Type": "application/x-www-form-urlencoded"})
    try:
        with urllib.request.urlopen(req, context=CTX, timeout=30) as r:
            return r.status, json.load(r)
    except urllib.error.HTTPError as e:
        try:
            return e.code, json.load(e)
        except Exception:
            return e.code, {"error": "http_" + str(e.code), "raw": e.read().decode("utf-8", "replace")}


def graph_get(url, token):
    req = urllib.request.Request(url, headers={"Authorization": "Bearer " + token})
    with urllib.request.urlopen(req, context=CTX, timeout=30) as r:
        return json.load(r)


def graph_post(url, token, payload):
    body = json.dumps(payload).encode()
    req = urllib.request.Request(url, body, {
        "Authorization": "Bearer " + token, "Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req, context=CTX, timeout=30) as r:
            return r.status, ""
    except urllib.error.HTTPError as e:
        return e.code, e.read().decode("utf-8", "replace")


def main():
    print("=" * 60)
    print("  Test: Geraetecode-Anmeldung + eine Testmail")
    print("=" * 60)

    # 1) Geraetecode anfordern
    status, dc = post_form(BASE + "/devicecode", {"client_id": CLIENT_ID, "scope": SCOPE})
    if "device_code" not in dc:
        print("\nFEHLGESCHLAGEN schon beim Start (Tenant erlaubt diesen Login nicht):")
        print("  ", dc.get("error"), "-", dc.get("error_description", "")[:300])
        return

    print("\n>>> SO ANMELDEN:")
    print("   ", dc.get("message", "Gehe auf https://microsoft.com/devicelogin und gib den Code ein."))
    print("\n(Warte auf deine Anmeldung im Browser ...)\n")

    # 2) Auf Anmeldung warten
    interval = int(dc.get("interval", 5))
    deadline = time.time() + int(dc.get("expires_in", 900))
    token = None
    while time.time() < deadline:
        time.sleep(interval)
        st, tok = post_form(BASE + "/token", {
            "grant_type": "urn:ietf:params:oauth:grant-type:device_code",
            "client_id": CLIENT_ID, "device_code": dc["device_code"]})
        if "access_token" in tok:
            token = tok["access_token"]
            break
        err = tok.get("error")
        if err in ("authorization_pending", "slow_down"):
            if err == "slow_down":
                interval += 5
            continue
        print("\nANMELDUNG NICHT MOEGLICH (Tenant blockt es):")
        print("  ", err, "-", tok.get("error_description", "")[:400])
        return

    if not token:
        print("\nZeit abgelaufen. Bitte erneut starten.")
        return

    # 3) Wer bin ich?
    try:
        me = graph_get("https://graph.microsoft.com/v1.0/me?$select=displayName,mail,userPrincipalName", token)
        addr = me.get("mail") or me.get("userPrincipalName")
        print("\nAngemeldet als:", me.get("displayName"), "<" + str(addr) + ">")
    except Exception as e:
        print("\nToken erhalten, aber /me schlug fehl:", e)
        return

    # 4) Eine Testmail an mich selbst
    payload = {
        "message": {
            "subject": "[TEST] ElternTermine – Graph funktioniert",
            "body": {"contentType": "HTML",
                     "content": "<div>Dies ist eine <b>Testmail</b> aus dem ElternTermine-Test. "
                                "Wenn du sie bekommst, funktioniert der Weg ohne IT.</div>"},
            "toRecipients": [{"emailAddress": {"address": addr}}]
        },
        "saveToSentItems": True
    }
    st, body = graph_post("https://graph.microsoft.com/v1.0/me/sendMail", token, payload)
    if st == 202:
        print("\n=========================================")
        print("  ERFOLG! Testmail an", addr, "gesendet.")
        print("  -> Dieser Weg funktioniert OHNE IT. Sag Claude Bescheid.")
        print("=========================================")
    else:
        print("\nSenden abgelehnt (HTTP " + str(st) + "):")
        print("  ", body[:500])
        print("\n-> Mail.Send ist in eurem Tenant gesperrt. Sag Claude den Fehlertext.")


if __name__ == "__main__":
    main()
