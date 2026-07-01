# IT-Rollout: Add-in "Mail an Todoist"

## Anfrage

Bitte das Add-in "Mail an Todoist" zentral bereitstellen über:
M365 Admin-Center > Einstellungen > Integrierte Apps > Benutzerdefinierte Apps hochladen

Genau wie das bestehende "CMI Mail"-Add-in.

---

## Manifest-URL

```
https://manuelweingartner.github.io/outlook-todoist-addin/manifest.prod.xml
```

---

## Benötigte Berechtigungen

- NUR Office.js `ReadItem` (im Manifest deklariert unter `<Permissions>ReadItem</Permissions>`).
- KEIN Microsoft Graph.
- KEIN `Mail.Read`-Consent.
- KEIN Azure-App-Registration.
- KEIN Mailbox-Lese-Consent auf Admin-Ebene.

Das Add-in läuft vollständig client-seitig im Office-Kontext. Es braucht keine
delegierten oder Anwendungsberechtigungen in Azure AD.

---

## Datenfluss

1. Nutzer öffnet eine Mail in Outlook und klickt "An Todoist-Task".
2. Das Add-in liest die aktuell geöffnete Mail lokal im Outlook-Client (Office.js,
   kein Netzwerk zu Microsoft).
3. Die Mail wird client-seitig als .eml-Datei aufgebaut.
4. Der Nutzer gibt seinen persönlichen Todoist-API-Token ein (einmalig, gespeichert
   in Office Roaming Settings).
5. Die .eml-Datei wird direkt vom Browser zu `api.todoist.com` hochgeladen.
6. Es wird kein Drittanbieter-Server ausser Todoist kontaktiert.
7. Jeder Nutzer verwendet seinen eigenen persönlichen Todoist-Account.

---

## Zuweisung

Bitte das Add-in zuweisen an:
- Manuel Weingartner (oder eine definierte Gruppe nach Absprache).

---

## Technische Details

- Hosting: GitHub Pages (statische Dateien, kein Server).
- Manifest-Requirement: Mailbox API 1.8 (unterstützt in Outlook 2019, M365).
- Kompatibilität: Outlook im Web, Outlook für Windows/Mac (modern).
