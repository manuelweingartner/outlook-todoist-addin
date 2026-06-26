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
- **Todoist direkt aus dem Browser** mit persoenlichem API-Token (Bearer). CORS erlaubt
  das (`Access-Control-Allow-Origin: *`), daher kein Proxy. Credentials-Mode bleibt
  `omit` (keine Cookies).
- Token wird in `Office.context.roamingSettings` gespeichert (roamt pro Nutzer).

## Module

- `src/lib/settings.ts` - Token lesen/speichern (roamingSettings).
- `src/lib/emlBuilder.ts` - `MailData` -> gueltiges MIME (.eml). Reine Logik.
- `src/lib/mailReader.ts` - Office.js-Item -> `MailData` (nur File-Anhaenge).
- `src/lib/todoist.ts` - REST v2 + Sync v9: getTasks/searchTasks/uploadFile/addComment.
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

- 2026-06-26: Code fertig, getestet, reviewt, live deployt. IT-Mail (docs/IT-EMAIL.md)
  an die CMI-IT verschickt; **wartet auf zentralen Rollout**. Erster echter E2E-Test
  erfolgt auf dem CMI-Postfach NACH dem Rollout.
- **CMI-Tenant blockiert User-Sideloading** -> Add-in muss von IT zentral ausgerollt
  werden (Praezedenz: "CMI Mail"). Selbst-Sideload ist NICHT moeglich.
- M365-Developer-Sandbox wurde verweigert (Programm verschaerft) -> kein Self-Test-Tenant.
- Neues Outlook hat kaputten App-Drag (Ziehen direkt in Todoist scheitert) und kein
  COM/Scripting -> deshalb das Add-in.
- Spec + Plan + SDD-Ledger: `docs/superpowers/`.
