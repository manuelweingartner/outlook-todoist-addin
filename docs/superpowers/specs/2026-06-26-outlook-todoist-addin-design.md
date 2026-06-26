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

Statisches Web-Add-in (Office.js), keine eigene Server-Logik. Die Mail wird **rein über
Office.js** ausgelesen (kein Microsoft Graph, kein Mail.Read-Consent); nur die Todoist-API
wird extern angesprochen — direkt aus dem Task-Pane-Browser-Kontext.

```
[Neues Outlook] --Office.js--> [Task-Pane (GitHub Pages, HTTPS)]
                                   |  1. Body + Anhänge via Office.js lesen,
                                   |     .eml client-seitig zusammenbauen
                                   v
                              [Task-Pane]
                                   |  2. Upload .eml + 3. Kommentar an Task
                                   v
                              [Todoist API]  (Auth: persönlicher API-Token, CORS *)
```

### Komponenten

| Komponente | Zweck | Abhängigkeiten |
|---|---|---|
| **Task-Pane Web-App** | UI: Task-Liste, Suche, Anhängen-Logik | Office.js |
| **Manifest** (XML/Unified) | Definiert Button + Pane; von IT ausgerollt | — |
| **EML-Builder-Modul** | Body + Anhänge → .eml (MIME) zusammenbauen | Office.js (ReadItem) |
| **Todoist-Modul** | Upload + Kommentar erstellen | Todoist-API-Token |
| **Settings-Modul** | Token speichern/laden | Office.roamingSettings |

Jedes Modul ist isoliert testbar (klare Funktionsgrenzen, gemockte HTTP-Aufrufe).

## Detaildesign

### Mail-Extraktion (nur Office.js, kein Graph)

- Header/Metadaten direkt aus `Office.context.mailbox.item`: `subject`, `from`, `to`, `cc`,
  `dateTimeCreated`, `normalizedSubject`.
- Body via `item.body.getAsync(Office.CoercionType.Html)` (HTML) — optional zusätzlich
  Plaintext für einen `multipart/alternative`-Teil.
- Datei-Anhänge: `item.attachments` durchgehen, je Anhang
  `item.getAttachmentContentAsync(id)` → Base64-Inhalt (`AttachmentContentFormat.Base64`).
- **EML-Builder** setzt daraus ein gültiges MIME-Dokument zusammen
  (`multipart/mixed`: Header + HTML-Body-Teil + je Anhang ein Base64-Teil) → `.eml`,
  die in Outlook geöffnet werden kann.
- Benötigt **nur** die Standard-`ReadItem`-Berechtigung des Add-ins (im Manifest, wird
  beim Installieren erteilt) — **kein Admin-Consent, kein Azure-App, kein Graph**.
- Bekannte Grenzen (akzeptiert): Original-Empfangs-Header (`Received:` etc.) fehlen;
  Inline-/eingebettete Bilder evtl. nicht 1:1; `item.attachments` listet keine
  Inline-only-Cloud-Attachments. Für Archiv/Referenz am Task ausreichend.

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
- `getAttachmentContentAsync` schlägt fehl / Anhang nicht lesbar → Anhang überspringen,
  Nutzer melden welche fehlen (statt still verwerfen).
- .eml > 25 MB (Todoist-Upload-Limit) → Hinweis + Abbruch.
- Doppel-Anhängen verhindern → Button während Upload sperren, Erfolg/Fehler eindeutig melden.
- **Keine stillen `try/catch`** — jeder Fehler wird sichtbar geloggt und dem Nutzer gemeldet
  (gemäss Defensive-Catch-Antipattern-Regel).

## Tests

- **Unit:** EML-Builder (Header/Body/Anhänge → gültiges MIME), Upload-Payload-Bau,
  Kommentar-Payload-Bau, Token-Persistenz (alle Office.js-/HTTP-Aufrufe gemockt).
- **Manuell/Integration:** Sideload in einem **Microsoft-365-Developer-Test-Tenant**
  (dort darf man selbst sideloaden) → End-to-End gegen ein Test-Postfach + Test-Todoist,
  bevor die IT den Produktiv-Rollout macht.

## Deployment

Statische Dateien auf **GitHub Pages** (Repo: privat oder öffentlich, Manifest-URL stabil).

**IT-Anfrage an CMI (konkret):**
1. Custom-Add-in zentral bereitstellen (Manifest-URL), wie „CMI Mail" — **das ist alles.**
   Kein Mailbox-Lese-Consent, keine Azure-App, kein Security-Review für Graph-Scopes.

## Offene Punkte / Risiken

- **Showstopper reduziert:** IT muss nur das Add-in deployen (Präzedenz „CMI Mail"
  vorhanden) → deutlich wahrscheinlichere/schnellere Freigabe.
- Manifest-Format: Unified (JSON) vs. klassisch (XML) — im Plan zu entscheiden; klassisches
  XML-Manifest ist für Outlook-Add-in-Deployment am breitesten unterstützt.
- EML-Fidelity der Rekonstruktion in der Praxis prüfen; falls unzureichend, ist Graph
  (mit Mail.Read-Consent) ein späteres optionales Upgrade — nicht jetzt.

## Nicht im Scope (YAGNI)

- Neue Tasks erstellen (macht das offizielle Todoist-Add-in bereits).
- OAuth-Flow für Todoist (Token reicht).
- Eigenes Backend/Proxy (dank CORS nicht nötig).
- Multi-User-Mandantenfähigkeit über Manuel hinaus (jeder Nutzer setzt eigenen Token).
