# IT-Rollout: Add-in "Mail an Todoist"

## Update 1.1.0.0 (November 2026): Outlook auf iPhone und Android

Das Manifest hat sich geändert (Version 1.0.0.0 -> 1.1.0.0). Bitte im M365 Admin-Center
unter Einstellungen > Integrierte Apps > "Mail an Todoist" die App aktualisieren bzw. das
neue Manifest von der gleichen URL hochladen. Neu:

- Mobile-Teil (`MobileFormFactor`), damit der Button auch in Outlook für iOS und Android
  erscheint (nur beim Lesen einer Mail).
- Mindest-API von Mailbox 1.8 auf 1.5 gesenkt (Outlook mobile unterstützt offiziell 1.5).
- **Berechtigungen unverändert:** weiterhin nur `ReadItem`, kein Graph, kein Consent.
- Zuweisung unverändert. Bis der Button auf den Geräten erscheint, kann es bis zu 24 h dauern.

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
- Manifest-Requirement: Mailbox API 1.5 (ab Version 1.1.0.0; Anhänge nutzen auf Desktop/Web 1.8 per Feature-Check).
- Kompatibilität: Outlook im Web, Outlook für Windows/Mac (modern), Outlook für iOS und Android (ab 1.1.0.0, dort ohne Anhänge).
