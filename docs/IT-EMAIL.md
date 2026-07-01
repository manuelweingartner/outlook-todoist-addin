# IT-Mail (Entwurf): Bitte um zentrale Bereitstellung

> Anrede/Empfänger und Ton vor dem Senden anpassen.

**Betreff:** Custom Outlook-Add-in "Mail an Todoist" - Bitte um zentrale Bereitstellung

---

Hallo [IT-Team / Name],

ich habe ein kleines Outlook-Add-in entwickelt, das die aktuell geöffnete Mail als `.eml` an eine bestehende Todoist-Aufgabe anhängt. Da das Sideloading eigener Add-ins im Tenant gesperrt ist, bitte ich euch um die **zentrale Bereitstellung** - genau wie beim bestehenden Add-in "CMI Mail".

**Bereitstellung:** M365 Admin-Center > Einstellungen > Integrierte Apps > Benutzerdefinierte Apps hochladen

**Manifest-URL:**
`https://manuelweingartner.github.io/outlook-todoist-addin/manifest.prod.xml`

**Benötigte Berechtigung - bewusst minimal:**
- Nur Office.js **`ReadItem`** (im Manifest deklariert).
- **Kein** Microsoft Graph, **kein** `Mail.Read`-Consent, **keine** Azure-App-Registrierung, **kein** Mailbox-Lese-Consent auf Admin-Ebene.

**Datenfluss (Datenschutz):** Das Add-in läuft vollständig client-seitig. Es liest nur die gerade geöffnete Mail im Outlook-Client, baut daraus lokal eine `.eml` und lädt diese zum persönlichen Todoist-Konto des Nutzers (`api.todoist.com`) hoch. Es wird kein Drittanbieter-Server ausser Todoist kontaktiert; jeder Nutzer verwendet seinen eigenen Todoist-Token.

**Zuweisung:** vorerst nur an mich (Manuel Weingartner); bei Interesse später gerne eine definierte Gruppe.

Technische Eckdaten und der vollständige Datenfluss sind hier dokumentiert: https://github.com/manuelweingartner/outlook-todoist-addin (siehe `docs/IT-ROLLOUT.md`). Bei Fragen oder für ein kurzes Review stehe ich gerne bereit.

Besten Dank und Gruss
Manuel
