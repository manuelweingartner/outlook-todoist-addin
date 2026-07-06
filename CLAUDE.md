# CLAUDE.md - Outlook-Add-in "Mail an Todoist"

Projektkontext für Claude-Code-Sessions. Halte das hier aktuell.

## Was es ist

Outlook-Web-Add-in (Office.js, TypeScript), das die aktuell geöffnete Mail als `.eml`
an einen per Klick gewählten **bestehenden** Todoist-Task anhängt. Löst: heute müsste
man die Mail per Drag auf den Desktop und dann in den Task ziehen; das offizielle
Todoist-Add-in kann nur neue Tasks erstellen, nicht an bestehende anhängen.

Ablauf für den Nutzer: Mail öffnen > Button "An Todoist-Task" > Task-Pane > Task
anklicken > .eml hängt als Kommentar-Anhang am Task.

## Architektur (bewusste Entscheidungen)

- **Reines statisches Add-in, kein Backend.** Gehostet auf GitHub Pages.
- **Mail-Extraktion nur über Office.js** (Body + Anhänge), `.eml` wird client-seitig
  zusammengebaut. **KEIN Microsoft Graph, KEIN Mail.Read-Consent, KEINE Azure-App.**
  Manifest-Permission ist nur `ReadItem`. (Das minimiert die IT-Freigabe; Graph wäre
  ein späteres optionales Upgrade für perfektere .eml-Fidelity.)
- **Todoist direkt aus dem Browser** mit persönlichem API-Token (Bearer). Die v1-API
  schickt für unsere Origin korrekte CORS-Header (`Access-Control-Allow-Origin` spiegelt
  `manuelweingartner.github.io`, `Allow-Headers: Authorization,Content-Type`), daher kein
  Proxy. Credentials-Mode bleibt `omit` (keine Cookies).
- Token wird in `Office.context.roamingSettings` gespeichert (roamt pro Nutzer).

## Module

- `src/lib/settings.ts` - Token lesen/speichern (roamingSettings).
- `src/lib/emlBuilder.ts` - `MailData` -> gültiges MIME (.eml). Reine Logik.
- `src/lib/mailReader.ts` - Office.js-Item -> `MailData` (nur File-Anhänge).
- `src/lib/todoist.ts` - **Unified API v1** (`api.todoist.com/api/v1`): `getAllTasks` (Cursor-Pagination über `/tasks?limit=200`, folgt `next_cursor` bis zum Ende, `MAX_PAGES`-Sicherung gegen Endlosschleifen bei instabilem Cursor, Basis für Client-Suche + Vorschläge), uploadFile (`/uploads`), addComment (`/comments`, gibt Kommentar-id zurück), deleteComment, getProjects, `createTask` (nimmt `NewTaskOptions`: content/project_id/priority/due_date/due_string/due_lang), `isAuthError` (401/403-Erkennung fürs Token-Retry). `TodoistTask` trägt optional `priority`/`due`. (Die frühere serverseitige Suche `getTasks`/`searchTasks`/`tasksByQuery` wurde entfernt, siehe Status 2026-07-06.)
- `src/lib/attachToTask.ts` - `senderName` (parst "Name <adresse>" fürs Kommentar-Meta), `prepareMail(mail, includeAttachments)` (baut .eml wahlweise mit/ohne Anhänge, liefert sizeBytes/subject/commentText/bodyText), `readAndPrepareCurrentMail` (liest die offene Mail einmal und liefert `MailData` + `PreparedMail` zusammen, Basis für den Text-only-Fallback im Task-Pane), `prepareCurrentMail` (Kompatibilitäts-Wrapper), `attachPreparedToTask` (Vorab-25MB-Check gegen `MAX_BYTES`, upload, Kommentar, gibt id), `formatMailDate` (UTC-stabil). Kommentartext = Betreff (Datum, von Name).
- `src/taskpane/taskLogic.ts` - Reine Logik: groupTasks (Heute/Überfällig), priorityColor, taskDeepLink, todayIso, filterTasks (case-insensitive Client-Suche, UND-Verknüpfung über Titel + Projektname, `#projekt`-Wörter matchen nur den Projektnamen), dueTodayOrOverdue (repliziert den früheren Server-Filter client-seitig), extractMailKeywords/suggestTasks (Top-3-Vorschläge aus Betreff-/Body-Schlagwörtern der aktuellen Mail), moveSelection (Pfeiltasten-Wrap für die Trefferliste), buildNewTaskOptions (validiert/mappt das Neuer-Task-Formular auf `NewTaskOptions`).
- `src/taskpane/taskpane.{html,ts,css}` - UI im Todoist-Look (rot `#e44332`, Dark-Mode via prefers-color-scheme, self-contained CSS, kein Fluent-CDN). 5 Zustände (Onboarding/Skeleton/Liste/Leer/Inline-Anhängen). Features: Ein-Klick-Anhängen mit Inline-Haken + Rückgängig, Mail-Kontext-Kopf, Vorab-Grössenwarnung, Tastatur-Flow (Enter hängt obersten Treffer an, Pfeiltasten navigieren die Trefferliste), Projektnamen, In-Todoist-öffnen, Neuen-Task-aus-Mail-Formular (Titel/Projekt/Priorität/Fälligkeit), Vorschläge-Sektion (passende Tasks zur offenen Mail zuoberst) + Client-Suche (case-insensitiv über alle geladenen Tasks, `#projekt`-Syntax), Retry-Banner bei fehlgeschlagenem Task-Load, Text-only-Anhängen-Button in der Grössenwarnung (Mail ohne Anhänge trotzdem anhängen, wenn die Vollversion über 25 MB liegt). Trefferzeile zweizeilig: Titel als prominente Hauptzeile (`.task-main`), Projekt als kleine Unterzeile.
- `src/open-task.html` - Statische Umleitungsseite (per Webpack-Copy ins dist-Root): läuft im System-Browser, löst dort `todoist://task?id=` aus (Desktop-Client), Web-App-Link als Fallback, schliesst sich nach 1.5s selbst. Nötig, weil die Outlook-Webview Custom-Protokolle blockt.
- `manifest.xml` (Dev, localhost:3000) / `manifest.prod.xml` (Prod, Pages-URL). **Redesign änderte das Manifest NICHT -> kein IT-Rollout nötig, Self-Deploy genügt.**

## Befehle

- Tests: `npx jest` (85 Tests, ts-jest 29, jsdom).
- Build: `npm run build` (Production-Webpack -> `dist/`).
- Manifest prüfen: `npx office-addin-manifest validate manifest.prod.xml`.
- **Deploy: `npm run deploy`** (baut + setzt `.nojekyll` + published `dist/` nach `gh-pages` via gh-pages-Tool).

## Deploy / Hosting

- Repo (public): https://github.com/manuelweingartner/outlook-todoist-addin
- Live (Pages, Branch `gh-pages`): https://manuelweingartner.github.io/outlook-todoist-addin/
- Manifest für IT: https://manuelweingartner.github.io/outlook-todoist-addin/manifest.prod.xml
- **Deploy ist manuell** (`npm run deploy`). Der CI-Auto-Deploy-Workflow wurde entfernt,
  weil der lokale `gh`-Token keinen `workflow`-Scope hatte. Zum Nachrüsten: Token den
  `workflow`-Scope geben (`gh auth refresh -s workflow`, richtigen Account wählen!) und
  `.github/workflows/deploy.yml` wieder anlegen.

## Konventionen (hart)

- Git-Identität: **Manuel Weingartner <manuel.weingartner@gmx.ch>**. NIE CMI. NIE
  Co-Authored-By Claude/Anthropic.
- **Keine Em-/En-Dashes** (`-` / `:` / `.` stattdessen).
- **Echte Umlaute** (ä/ö/ü) in sichtbarem Deutsch, nie ae/oe/ue. Schweizer Deutsch:
  kein ß (gross/ausser/Gruss/-mässig bleiben mit ss).
- Kein stilles `try/catch`: Fehler immer loggen UND dem Nutzer im Pane zeigen.
- Neue Logik kommt mit Test (TDD).

## Status / Gotchas

- 2026-07-06 spät (2): **Deep-Link-Fix (`35f3fa5`, deployed): todoist://task?id= braucht die ALTE numerische Id.**
  Symptom: Desktop-App öffnete, sprang aber nicht zum Task. Die v1-API liefert neue
  alphanumerische Ids, das URL-Scheme des Desktop-Clients versteht nur die alten
  numerischen (Doku-Beispiele durchweg numerisch; Übersetzung via
  `GET /api/v1/id_mappings/tasks/<ids>` -> `[{old_id, new_id}]`, gechunkt à 100).
  Umsetzung: `getOldTaskIds` (todoist.ts) lädt das Mapping nicht-blockierend nach dem
  Task-Load, `openInTodoist` (taskpane.ts) nutzt die alte Id, `taskDeepLink(appId, webId)`
  gibt beide weiter, `open-task.html` nimmt `id` fürs Protokoll und `web` für den
  Browser-Fallback. Ohne Mapping (Fehler/leer) Fallback aufs bisherige Verhalten.
  **GOTCHA 3: todoist://-Schemes (task/project) erwarten die alten numerischen Ids,
  nie die neuen alphanumerischen der v1-API. Immer über id_mappings übersetzen.**
  Offen: On-Device-Verifikation durch Manuel (Klick auf ↗ muss den Task öffnen).
- 2026-07-06 spät: **On-Device-Nachfix-Runde (3 Hotfixes direkt auf master, alle deployed + von Manuel abgenommen).**
  (1) `041daa7`: todoist://-Direktlinks revertiert (siehe Gotcha unten), Projekt-Dropdown
  lädt Projekte beim Formular-Öffnen nach falls Cache leer + "Inbox"-Fallback-Option,
  Retry-Button kompakt statt vollbreit. (2) `0570f46`: `color-scheme: light`/`dark` im CSS.
  (3) `a6fd65e`: Task-Links zeigen auf die Umleitungsseite `open-task.html?id=` -> öffnet
  den **Todoist-Desktop-Client** via System-Browser (Windows fragt einmal "Immer erlauben").
  **GOTCHA 1: Die Outlook-Webview blockt Custom-Protokoll-Links (todoist:// zeigt eine
  Verbotssymbol-Fehlerseite), auch via Anchor-Click. Protokoll-Sprünge IMMER über eine
  im System-Browser geöffnete Umleitungsseite lösen.**
  **GOTCHA 2: WebView2 braucht `color-scheme` im CSS, sonst rendert das native
  Select-Popup im Dark-Mode hell mit heller Textfarbe = Dropdown wirkt leer.**
- 2026-07-06: **Verbesserungspaket (5 Features, Branch `feat/verbesserungspaket`).**
  (1) Pfeiltasten-Navigation in der Trefferliste (`moveSelection`, wrapt an den Rändern,
  Enter hängt die aktuell markierte Zeile an, nicht mehr fix die oberste). (2) Neues-Task-
  Formular direkt aus der Mail (Titel/Projekt/Priorität/Fälligkeits-Chips, `buildNewTaskOptions`
  validiert, erstellt Task + hängt Mail in einem Zug an). (3) Ablage-Kommentar trägt jetzt
  den Absendernamen (`senderName`, Format "Betreff (Datum, von Name)"). (4) Client-Suche
  versteht `#projekt`-Syntax (matcht nur den Projektnamen, normale Wörter weiter Titel ODER
  Projekt). (5) Robustheit: `getAllTasks` hat eine `MAX_PAGES`-Sicherung gegen Endlosschleifen,
  Retry-Banner bei fehlgeschlagenem Task-Load, und ein Text-only-Fallback für übergrosse
  Mails (`readAndPrepareCurrentMail`/`prepareMail` trennen Mail-Rohdaten von der gebauten
  .eml, die Grössenwarnung bietet bei >25 MB einen "Ohne Anhänge anhängen"-Button, der die
  Mail ohne Anhänge neu baut; ist der Body allein schon zu gross, bleibt Anhängen deaktiviert
  mit entsprechendem Hinweis). 85 Tests grün, Build clean, `grep "25 \* 1024"` in
  `src/taskpane/taskpane.ts` liefert keine Treffer mehr (nur noch `MAX_BYTES`).
- 2026-07-06: **Tote Server-Suche entfernt (Cleanup nach Suche+Vorschläge-Feature).**
  `getTasks`/`searchTasks`/`tasksByQuery` aus `src/lib/todoist.ts` gelöscht, samt Tests.
  Was war/ist der Stand: Task-Pane nutzt seit dem Suche-und-Vorschläge-Feature nur noch
  `getAllTasks` (Cursor-Pagination, lädt jetzt ALLE Seiten statt vorher nur Seite 1),
  case-insensitive Client-Suche (`filterTasks`) und Top-3-Vorschläge (`suggestTasks`)
  aus Betreff-/Body-Schlagwörtern der offenen Mail. Warum entfernt: Todoists
  Server-Suche (`/tasks/filter?query=`) matchte case-sensitiv, das ist der Grund
  wieso überhaupt auf Client-Suche umgestellt wurde. Der "nacktes Array"-Defensivfall
  wurde in die `getAllTasks`-Suite migriert (dort war er noch nicht abgedeckt), der
  401-Fall existierte dort schon. 59 Tests grün, Build clean, kein Treffer mehr für
  `getTasks`/`searchTasks` in `src/`.
- 2026-06-30: **UX-Feinschliff (live deployed).** (1) Alle sichtbaren Task-Pane-Texte
  auf echte Umlaute umgestellt (Anhängen/hängen/öffnen/einfügen/für/Überfällig/Rückgängig/ungültig;
  `gross` bleibt Schweizer Deutsch, kein ß). (2) Trefferzeile umgebaut: Titel ist jetzt
  die prominente Hauptzeile (`.task-main` Spalten-Flex, `font-weight 500`, bis 2 Zeilen),
  Projekt als kleine graue Unterzeile darunter, Öffnen-Pfeil rechts (`li` ist jetzt Flex).
  39 Tests grün, Build clean, `npm run deploy` raus, Manuel hat Optik abgenommen. Alle
  Repo-Docs (.md) ebenfalls auf echte Umlaute normalisiert. **manifest.prod.xml-Tooltips
  bewusst NICHT angefasst** (Umlaut-Fix dort würde Manifest ändern -> potenzieller IT-Re-Rollout).
- 2026-06-29: **Task-Pane-Redesign** auf Branch `feat/taskpane-redesign` (Spec+Plan in
  `docs/superpowers/`). Todoist-Look, Dark-Mode, 5 Zustände, 9 UX-Features. 39 Tests grün,
  Build clean, **Manifest unverändert** (Self-Deploy genügt, kein IT-Rollout). OFFEN vor
  Live: (1) v1-API-Feldnamen `priority`/`due.date` mit echtem Token verifizieren (Mapping in
  taskLogic/todoist hängt dran), (2) Optik-Abnahme Light+Dark, (3) Deeplink
  `app.todoist.com/app/task/<id>` am Konto prüfen, (4) Deploy mit gh-Account-Switch + on-device-E2E.
- 2026-06-29: **Zentraler Rollout erfolgt** (Hadi/Interne IT, Ticket #140204, für Manuel
  freigeschaltet via M365 Integrierte Apps). Erster echter Aufruf gegen Todoist scheiterte
  mit `Failed to fetch`.
- **GOTCHA (Root Cause des Failed-to-fetch): Todoist hat REST v2 + Sync v9 abgeschaltet
  (HTTP 410 Gone).** Die 410-Antwort trägt KEINE CORS-Header -> der Browser blockt sie
  und meldet `Failed to fetch` (sieht aus wie Netz/Token-Problem, ist aber tote API-Version).
  Diagnose-Trick: `curl -i -X OPTIONS .../rest/v2/tasks -H "Origin: ..."` zeigt 410 ohne
  Access-Control-Header. Fix: Migration auf Unified **API v1** (`/api/v1`, commit-getestet,
  26 Tests). v1 liefert korrekte CORS-Header.
- 2026-06-26: Code fertig, getestet, reviewt, live deployt. IT-Mail (docs/IT-EMAIL.md)
  an die CMI-IT verschickt.
- **Deploy-Gotcha:** `npm run deploy` pusht via `gh auth git-credential`. Der AKTIVE
  gh-Account ist standardmässig `CMI-Kunden` (kein Zugriff aufs private Repo -> 403).
  Vor dem Deploy `gh auth switch --user manuelweingartner`, danach zurück auf `CMI-Kunden`.
  (Vor dem Deploy `gh auth status` prüfen; am 2026-06-30 war manuelweingartner schon aktiv.)
- **CMI-Tenant blockiert User-Sideloading** -> Add-in muss von IT zentral ausgerollt
  werden (Präzedenz: "CMI Mail"). Selbst-Sideload ist NICHT möglich.
- M365-Developer-Sandbox wurde verweigert (Programm verschärft) -> kein Self-Test-Tenant.
- Neues Outlook hat kaputten App-Drag (Ziehen direkt in Todoist scheitert) und kein
  COM/Scripting -> deshalb das Add-in.
- Spec + Plan + SDD-Ledger: `docs/superpowers/`.
