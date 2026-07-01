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
- `src/lib/todoist.ts` - **Unified API v1** (`api.todoist.com/api/v1`): getTasks/searchTasks (`/tasks/filter?query=`, paginiert -> `{results}`), uploadFile (`/uploads`), addComment (`/comments`, gibt Kommentar-id zurück), deleteComment, getProjects, createTask. `TodoistTask` trägt optional `priority`/`due`.
- `src/lib/attachToTask.ts` - `prepareCurrentMail` (liest+baut .eml einmal, liefert sizeBytes/subject/commentText), `attachPreparedToTask` (Vorab-25MB-Check, upload, Kommentar, gibt id), `formatMailDate` (UTC-stabil). Kommentartext = Betreff (Datum).
- `src/taskpane/taskLogic.ts` - Reine Logik: groupTasks (Heute/Überfällig), priorityColor, taskDeepLink, todayIso.
- `src/taskpane/taskpane.{html,ts,css}` - UI im Todoist-Look (rot `#e44332`, Dark-Mode via prefers-color-scheme, self-contained CSS, kein Fluent-CDN). 5 Zustände (Onboarding/Skeleton/Liste/Leer/Inline-Anhängen). Features: Ein-Klick-Anhängen mit Inline-Haken + Rückgängig, Mail-Kontext-Kopf, Vorab-Grössenwarnung, Tastatur-Flow (Enter hängt obersten Treffer an), Projektnamen, In-Todoist-öffnen, Neuen-Task-aus-Mail. Trefferzeile zweizeilig: Titel als prominente Hauptzeile (`.task-main`), Projekt als kleine Unterzeile.
- `manifest.xml` (Dev, localhost:3000) / `manifest.prod.xml` (Prod, Pages-URL). **Redesign änderte das Manifest NICHT -> kein IT-Rollout nötig, Self-Deploy genügt.**

## Befehle

- Tests: `npx jest` (39 Tests, ts-jest 29, jsdom).
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
