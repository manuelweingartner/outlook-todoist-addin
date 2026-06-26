# Design: Outlook-Add-in „Mail an Todoist-Task"

Datum: 2026-06-26
Status: Genehmigt (Brainstorming abgeschlossen)
Autor: Manuel Weingartner

## Problem

Eingehende E-Mails sollen mit wenigen Klicks als Anhang an einen **bestehenden**
Todoist-Task gehängt werden. Heutiger Weg: Mail per Drag&Drop auf den Desktop, dann
die Datei in den Task ziehen (zwei Drags + Kontextwechsel). Das Todoist-Outlook-Add-in
kann nur neue Tasks erstellen, nicht an bestehende anhängen.

### Harte Rahmenbedingungen (im Brainstorming festgestellt)

- **Neues Outlook für Windows** (Web-App-Kern): kein COM/Scripting, kaputter App-Drag
  (Ziehen direkt in Todoist scheitert), nur Drag ins Dateisystem funktioniert.
- **CMI-Tenant sperrt User-Sideloading** von Custom-Add-ins. Bestätigt: unter
  „Apps verwalten" gibt es nur „Weitere Apps erhalten", keine Upload-Option.
  Custom-Add-ins existieren aber zentral ausgerollt (z.B. „CMI Mail") → IT kann deployen.
- **Todoist Pro** vorhanden.
- Ziel-Task ist **jedes Mal ein anderer** → der Nutzer will den Task **anklicken**,
  nicht dessen Titel erinnern/tippen.

### Warum ein Add-in (und nicht die No-Code-Wege)

Alle Wege, die „neues Outlook" + „kein Sideloading" beibehalten (Forwarding an
Task-Adresse, Forward+Companion-Zuordnung), landen bei ~2 Aktionen + Kontextwechsel
und sind damit **nicht effizienter** als der heutige Doppel-Drag. Eine echt effiziente
1-Klick-Lösung mit Task-Liste direkt in Outlook erfordert ein Add-in. Da Self-Sideloading
gesperrt ist, wird es **von der CMI-IT zentral ausgerollt** (vom Nutzer als realistisch
eingeschätzt).

## Ziel-Ablauf (User Story)

1. Mail im neuen Outlook offen/markiert → Button **„An Todoist-Task"** in der Mail-Leiste.
2. Task-Pane öffnet sich rechts: Tasks (Heute/Überfällig) + Suchfeld + „zuletzt verwendet".
3. Ziel-Task **anklicken** → Add-in hängt die Original-Mail (.eml) als Kommentar-Anhang
   an den Task → kurze Erfolgsmeldung.

Drei Aktionen, kein Titel-Tippen, kein Drag, kein Datei-Umweg, alles in Outlook.

## Architektur

Statisches Web-Add-in (Office.js), keine eigene Server-Logik. Zwei externe APIs werden
direkt aus dem Task-Pane-Browser-Kontext angesprochen.

```
[Neues Outlook] --Office.js--> [Task-Pane (GitHub Pages, HTTPS)]
                                   |  1. Mail als MIME holen
                                   v
                              [Microsoft Graph]  GET /me/messages/{id}/$value
                                   |  (Auth: NAA/MSAL.js, delegiert Mail.Read)
                                   v
                              [Task-Pane]
                                   |  2. Upload .eml + 3. Kommentar an Task
                                   v
                              [Todoist API]  (Auth: persönlicher API-Token, CORS *)
```

### Komponenten

| Komponente | Zweck | Abhängigkeiten |
|---|---|---|
| **Task-Pane Web-App** | UI: Task-Liste, Suche, Anhängen-Logik | Office.js, MSAL.js |
| **Manifest** (XML/Unified) | Definiert Button + Pane; von IT ausgerollt | — |
| **Graph-Modul** | Holt Original-Mail als .eml | Azure-App + Mail.Read |
| **Todoist-Modul** | Upload + Kommentar erstellen | Todoist-API-Token |
| **Settings-Modul** | Token speichern/laden | Office.roamingSettings |

Jedes Modul ist isoliert testbar (klare Funktionsgrenzen, gemockte HTTP-Aufrufe).

## Detaildesign

### Mail-Extraktion (Microsoft Graph)

- `Office.context.mailbox.item.itemId` → `Office.context.mailbox.convertToRestId(...)`.
- `GET https://graph.microsoft.com/v1.0/me/messages/{restId}/$value` → volles MIME = .eml.
- Auth: **Nested App Authentication (NAA)** via MSAL.js → Graph-Token **client-seitig**,
  **kein Backend/OBO** nötig.
- Benötigt: **Azure-App-Registrierung** mit delegierter Berechtigung **`Mail.Read`** +
  **Admin-Consent** (Teil der IT-Anfrage).
- Bewusste Entscheidung gegen EWS (`makeEwsRequestAsync` + Callback-Token, consent-frei),
  da EWS für Exchange Online abgekündigt wird → nicht zukunftssicher.

### Todoist-Anbindung

- **Auth:** persönlicher API-Token (Todoist → Einstellungen → Integrationen → API-Token).
  Einmal im Pane eingefügt, gespeichert in `Office.context.roamingSettings`.
- **Anhängen (2 Schritte):**
  1. `POST /uploads` (multipart, die .eml-Datei) → Datei-Objekt.
  2. `POST /comments` mit `task_id` + Anhang-Objekt → Mail hängt als Kommentar am Task.
- CORS bestätigt: Todoist setzt `Access-Control-Allow-Origin: *` für authentifizierte
  Bearer-Token-Requests → Browser-Aufrufe ok, **kein Proxy nötig**. (Credentials-Mode
  bleibt `omit`/Standard; nur `Authorization`-Header, keine Cookies.)

### Task-Auswahl (UX)

- Default-Liste: **Heute + Überfällig** (`GET /tasks` mit Filter), kurz gehalten.
- **Suchfeld**: live gegen Todoist gefiltert; optionaler Projekt-Filter.
- **„Zuletzt verwendet"** oben (lokal in roamingSettings gemerkt).
- Klick auf einen Task = anhängen.

### Settings

- Token in `Office.context.roamingSettings` (roamt über Outlook-Installationen des Nutzers).
- Beim ersten Start oder bei 401: Token-Eingabe-Flow.

## Fehlerbehandlung

- Kein/ungültiger Todoist-Token (401) → Pane fordert Token (neu) an.
- Graph-Token/Consent fehlt → klare Meldung „IT-Freigabe ausstehend / Anmeldung nötig".
- .eml > 25 MB (Todoist-Upload-Limit) → Hinweis + Abbruch.
- Doppel-Anhängen verhindern → Button während Upload sperren, Erfolg/Fehler eindeutig melden.
- **Keine stillen `try/catch`** — jeder Fehler wird sichtbar geloggt und dem Nutzer gemeldet
  (gemäss Defensive-Catch-Antipattern-Regel).

## Tests

- **Unit:** Graph-MIME-Abruf, Upload-Payload-Bau, Kommentar-Payload-Bau, Token-Persistenz
  (alle HTTP-Aufrufe gemockt).
- **Manuell/Integration:** Sideload in einem **Microsoft-365-Developer-Test-Tenant**
  (dort darf man selbst sideloaden) → End-to-End gegen ein Test-Postfach + Test-Todoist,
  bevor die IT den Produktiv-Rollout macht.

## Deployment

Statische Dateien auf **GitHub Pages** (Repo: privat oder öffentlich, Manifest-URL stabil).

**IT-Anfrage an CMI (konkret):**
1. Custom-Add-in zentral bereitstellen (Manifest-URL), wie „CMI Mail".
2. Azure-App-Registrierung + **Admin-Consent für delegiert `Mail.Read`**.

## Offene Punkte / Risiken

- **Einziger echter Showstopper:** IT-Freigabe (Add-in-Deployment + `Mail.Read`-Consent).
  Vom Nutzer als realistisch eingeschätzt; kein Graph-Fallback eingeplant.
- Manifest-Format: Unified (JSON) vs. klassisch (XML) — im Plan zu entscheiden; klassisches
  XML-Manifest ist für Outlook-Add-in-Deployment am breitesten unterstützt.
- NAA-Verfügbarkeit im neuen Outlook für Windows verifizieren (sollte unterstützt sein).

## Nicht im Scope (YAGNI)

- Neue Tasks erstellen (macht das offizielle Todoist-Add-in bereits).
- OAuth-Flow für Todoist (Token reicht).
- Eigenes Backend/Proxy (dank CORS nicht nötig).
- Multi-User-Mandantenfähigkeit über Manuel hinaus (jeder Nutzer setzt eigenen Token).
