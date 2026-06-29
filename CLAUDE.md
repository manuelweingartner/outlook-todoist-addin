# CLAUDE.md - Outlook-Add-in "Mail an Todoist"

Projektkontext fuer Claude-Code-Sessions. Halte das hier aktuell.

## Was es ist

Outlook-Web-Add-in (Office.js, TypeScript), das die aktuell geoeffnete Mail als `.eml`
an einen per Klick gewaehlten **bestehenden** Todoist-Task anhaengt. Loest: heute muesste
man die Mail per Drag auf den Desktop und dann in den Task ziehen; das offizielle
Todoist-Add-in kann nur neue Tasks erstellen, nicht an bestehende anhaengen.

Ablauf fuer den Nutzer: Mail oeffnen > Button "An Todoist-Task" > Task-Pane > Task
anklicken > .eml haengt als Kommentar-Anhang am Task.

## Architektur (bewusste Entscheidungen)

- **Reines statisches Add-in, kein Backend.** Gehostet auf GitHub Pages.
- **Mail-Extraktion nur ueber Office.js** (Body + Anhaenge), `.eml` wird client-seitig
  zusammengebaut. **KEIN Microsoft Graph, KEIN Mail.Read-Consent, KEINE Azure-App.**
  Manifest-Permission ist nur `ReadItem`. (Das minimiert die IT-Freigabe; Graph waere
  ein spaeteres optionales Upgrade fuer perfektere .eml-Fidelity.)
- **Todoist direkt aus dem Browser** mit persoenlichem API-Token (Bearer). Die v1-API
  schickt fuer unsere Origin korrekte CORS-Header (`Access-Control-Allow-Origin` spiegelt
  `manuelweingartner.github.io`, `Allow-Headers: Authorization,Content-Type`), daher kein
  Proxy. Credentials-Mode bleibt `omit` (keine Cookies).
- Token wird in `Office.context.roamingSettings` gespeichert (roamt pro Nutzer).

## Module

- `src/lib/settings.ts` - Token lesen/speichern (roamingSettings).
- `src/lib/emlBuilder.ts` - `MailData` -> gueltiges MIME (.eml). Reine Logik.
- `src/lib/mailReader.ts` - Office.js-Item -> `MailData` (nur File-Anhaenge).
- `src/lib/todoist.ts` - **Unified API v1** (`api.todoist.com/api/v1`): getTasks/searchTasks (`/tasks/filter?query=`, paginiert -> `{results}`), uploadFile (`/uploads`), addComment (`/comments`).
- `src/lib/attachToTask.ts` - Orchestrierung: lesen -> .eml bauen -> 25MB-Check -> upload -> Kommentar.
- `src/taskpane/taskpane.{html,ts,css}` - UI: Token-Flow, Task-Liste (Heute/Ueberfaellig + Suche), Anhaengen mit Status.
- `manifest.xml` (Dev, localhost:3000) / `manifest.prod.xml` (Prod, Pages-URL).

## Befehle

- Tests: `npx jest` (25 Tests, ts-jest 29, jsdom).
- Build: `npm run build` (Production-Webpack -> `dist/`).
- Manifest pruefen: `npx office-addin-manifest validate manifest.prod.xml`.
- **Deploy: `npm run deploy`** (baut + setzt `.nojekyll` + published `dist/` nach `gh-pages` via gh-pages-Tool).

## Deploy / Hosting

- Repo (public): https://github.com/manuelweingartner/outlook-todoist-addin
- Live (Pages, Branch `gh-pages`): https://manuelweingartner.github.io/outlook-todoist-addin/
- Manifest fuer IT: https://manuelweingartner.github.io/outlook-todoist-addin/manifest.prod.xml
- **Deploy ist manuell** (`npm run deploy`). Der CI-Auto-Deploy-Workflow wurde entfernt,
  weil der lokale `gh`-Token keinen `workflow`-Scope hatte. Zum Nachruesten: Token den
  `workflow`-Scope geben (`gh auth refresh -s workflow`, richtigen Account waehlen!) und
  `.github/workflows/deploy.yml` wieder anlegen.

## Konventionen (hart)

- Git-Identitaet: **Manuel Weingartner <manuel.weingartner@gmx.ch>**. NIE CMI. NIE
  Co-Authored-By Claude/Anthropic.
- **Keine Em-/En-Dashes** (`-` / `:` / `.` stattdessen).
- Kein stilles `try/catch`: Fehler immer loggen UND dem Nutzer im Pane zeigen.
- Neue Logik kommt mit Test (TDD).

## Status / Gotchas

- 2026-06-29: **Zentraler Rollout erfolgt** (Hadi/Interne IT, Ticket #140204, fuer Manuel
  freigeschaltet via M365 Integrierte Apps). Erster echter Aufruf gegen Todoist scheiterte
  mit `Failed to fetch`.
- **GOTCHA (Root Cause des Failed-to-fetch): Todoist hat REST v2 + Sync v9 abgeschaltet
  (HTTP 410 Gone).** Die 410-Antwort traegt KEINE CORS-Header -> der Browser blockt sie
  und meldet `Failed to fetch` (sieht aus wie Netz/Token-Problem, ist aber tote API-Version).
  Diagnose-Trick: `curl -i -X OPTIONS .../rest/v2/tasks -H "Origin: ..."` zeigt 410 ohne
  Access-Control-Header. Fix: Migration auf Unified **API v1** (`/api/v1`, commit-getestet,
  26 Tests). v1 liefert korrekte CORS-Header.
- 2026-06-26: Code fertig, getestet, reviewt, live deployt. IT-Mail (docs/IT-EMAIL.md)
  an die CMI-IT verschickt.
- **Deploy-Gotcha:** `npm run deploy` pusht via `gh auth git-credential`. Der AKTIVE
  gh-Account ist standardmaessig `CMI-Kunden` (kein Zugriff aufs private Repo -> 403).
  Vor dem Deploy `gh auth switch --user manuelweingartner`, danach zurueck auf `CMI-Kunden`.
- **CMI-Tenant blockiert User-Sideloading** -> Add-in muss von IT zentral ausgerollt
  werden (Praezedenz: "CMI Mail"). Selbst-Sideload ist NICHT moeglich.
- M365-Developer-Sandbox wurde verweigert (Programm verschaerft) -> kein Self-Test-Tenant.
- Neues Outlook hat kaputten App-Drag (Ziehen direkt in Todoist scheitert) und kein
  COM/Scripting -> deshalb das Add-in.
- Spec + Plan + SDD-Ledger: `docs/superpowers/`.
