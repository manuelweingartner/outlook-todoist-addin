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

- `src/lib/settings.ts` - Token lesen/speichern (roamingSettings): `getToken`/`setToken` (Schlüssel `todoistToken`) + `getAnthropicKey`/`setAnthropicKey` (Schlüssel `anthropicKey`, optional, für die KI-Zusammenfassung; leerer String = aus).
- `src/lib/summarize.ts` - KI-Zusammenfassung des Mailinhalts als Kommentartitel. `buildSummaryPrompt`/`parseSummary` (rein, testbar) + `summarizeMail(key, mail, bodyText)`: raw `fetch` an `api.anthropic.com/v1/messages`, Modell `claude-haiku-4-5`, Header `anthropic-dangerous-direct-browser-access: true` (CORS für Browser erlaubt, verifiziert `allow-origin: *`). Wirft `SummaryError`; Aufrufer fällt auf den Betreff-Titel zurück.
- `src/lib/emlBuilder.ts` - `MailData` -> gültiges MIME (.eml). Reine Logik.
- `src/lib/mailReader.ts` - Office.js-Item -> `MailData` (nur File-Anhänge).
- `src/lib/todoist.ts` - **Unified API v1** (`api.todoist.com/api/v1`): `getAllTasks` (Cursor-Pagination über `/tasks?limit=200`, folgt `next_cursor` bis zum Ende, `MAX_PAGES`-Sicherung gegen Endlosschleifen bei instabilem Cursor, Basis für Client-Suche + Vorschläge), uploadFile (`/uploads`), addComment (`/comments`, gibt Kommentar-id zurück), deleteComment, getProjects, `createTask` (nimmt `NewTaskOptions`: content/project_id/priority/due_date/due_string/due_lang), `isAuthError` (401/403-Erkennung fürs Token-Retry). `TodoistTask` trägt optional `priority`/`due`. (Die frühere serverseitige Suche `getTasks`/`searchTasks`/`tasksByQuery` wurde entfernt, siehe Status 2026-07-06.)
- `src/lib/attachToTask.ts` - `senderName` (parst "Name <adresse>" fürs Kommentar-Meta), `prepareMail(mail, includeAttachments)` (baut .eml wahlweise mit/ohne Anhänge, liefert sizeBytes/subject/commentText/bodyText), `readAndPrepareCurrentMail` (liest die offene Mail einmal und liefert `MailData` + `PreparedMail` zusammen, Basis für den Text-only-Fallback im Task-Pane), `prepareCurrentMail` (Kompatibilitäts-Wrapper), `attachPreparedToTask` (Vorab-25MB-Check gegen `MAX_BYTES`, upload, Kommentar, gibt id), `formatMailDate` (UTC-stabil). Kommentartext = Betreff (Datum, von Name).
- `src/taskpane/taskLogic.ts` - Reine Logik: groupTasks (Heute/Überfällig), priorityColor, taskDeepLink, todayIso, filterTasks (case-insensitive Client-Suche, UND-Verknüpfung über Titel + Projektname, `#projekt`-Wörter matchen nur den Projektnamen), dueTodayOrOverdue (repliziert den früheren Server-Filter client-seitig), extractMailKeywords/suggestTasks (Top-3-Vorschläge aus Betreff-/Body-Schlagwörtern der aktuellen Mail), moveSelection (Pfeiltasten-Wrap für die Trefferliste), buildNewTaskOptions (validiert/mappt das Neuer-Task-Formular auf `NewTaskOptions`).
- `src/taskpane/taskpane.{html,ts,css}` - UI im Todoist-Look (rot `#e44332`, Dark-Mode via prefers-color-scheme, self-contained CSS, kein Fluent-CDN). 5 Zustände (Onboarding/Skeleton/Liste/Leer/Inline-Anhängen). Features: Ein-Klick-Anhängen mit Inline-Haken + Rückgängig, Mail-Kontext-Kopf, Vorab-Grössenwarnung, Tastatur-Flow (Enter hängt obersten Treffer an, Pfeiltasten navigieren die Trefferliste), Projektnamen, In-Todoist-öffnen, Neuen-Task-aus-Mail-Formular (Titel/Projekt/Priorität/Fälligkeit), Vorschläge-Sektion (passende Tasks zur offenen Mail zuoberst) + Client-Suche (case-insensitiv über alle geladenen Tasks, `#projekt`-Syntax), Retry-Banner bei fehlgeschlagenem Task-Load, Text-only-Anhängen-Button in der Grössenwarnung (Mail ohne Anhänge trotzdem anhängen, wenn die Vollversion über 25 MB liegt). Trefferzeile zweizeilig: Titel als prominente Hauptzeile (`.task-main`), Projekt als kleine Unterzeile.
- `src/open-task.html` - Statische Umleitungsseite (per Webpack-Copy ins dist-Root): läuft im System-Browser und löst dort das Desktop-Protokoll aus. Nötig, weil die Outlook-Webview Custom-Protokolle blockt. **Aktueller Stand (nach der Deeplink-Saga vom 2026-07-06, siehe Status):** feuert `todoist://navigate-to?url=https://app.todoist.com/app/task/<neue-id>` (NICHT `todoist://task?id=` -> das triggert die "Link aktualisieren"-Warnung), ruft KEIN `window.close()` auf (das kappte Chromes Erlaubnis-Dialog) und zeigt einen prominenten Klick-Button "Task in Todoist öffnen" (Nutzergeste öffnet zuverlässig) plus Web-App-Fallback-Link.
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

- 2026-07-07 (2): **"Sehe den Change nicht" = WebView2-Cache des NEUEN Outlook.** Manuel nutzt
  das neue Outlook (`olk.exe` + Add-in-Host `olkexthost.exe`). Dessen Add-in-WebView2-Cache liegt
  NICHT im klassischen `...\Office\16.0\Wef\`, sondern in
  **`%LOCALAPPDATA%\Microsoft\Olk\EBWebView\Default\Cache\Cache_Data`** (+ `Code Cache\js|wasm`).
  Weil `taskpane.js` frueher einen stabilen Namen hatte, servierte WebView2 die alte Version von
  dort - ein Outlook-Neustart leert diesen Cache NICHT. Zwei Fixes: (a) Sofort-Recovery: olk/
  olkexthost schliessen (neue-Outlook-Entwuerfe sind server-synchron, Schliessen sicher), dann
  NUR `Cache\Cache_Data` + `Code Cache\js|wasm` leeren (nicht das ganze EBWebView-Profil - das
  haelt Cookies/Login), reopen. (b) Dauerhaft: **Content-Hash im Prod-Bundle-Namen**
  (`output.filename: "[name].[contenthash].js"` in webpack.config.js) -> jede Aenderung ergibt
  eine neue JS-URL, HtmlWebpackPlugin injiziert sie in die fix benannte `taskpane.html`
  (Manifest-Ziel, daher kein IT-Re-Rollout). **GOTCHA: Add-in-Cache-Debugging beim neuen Outlook
  immer im Olk\EBWebView-Cache, nicht im Wef-Ordner. Bundle-Namen gehasht halten.**

- 2026-07-07: **KI-Zusammenfassung als Kommentartitel LIVE (opt-in, nicht-brechend).** Statt
  `Betreff (Datum, von Name)` erzeugt das Add-in beim Anhängen einen Ein-Satz-Zusammenfassung
  der Mail als Todoist-Kommentartitel (von wem/worum/welche Aktion), client-seitig via Claude
  Haiku 4.5. **Aktiv nur wenn ein Anthropic-Key hinterlegt ist** (zweites Feld im Einstellungs-
  bereich, Schlüssel `anthropicKey` in roamingSettings; "⚙ Einstellungen"-Button in der
  Task-Ansicht öffnet den Bereich vorbefüllt). Kein Key -> exakt wie bisher (Betreff-Titel).
  Fully-automatic (kein Vorschau-Schritt), bestehender Rückgängig-Button. Fehler-Fallback:
  Key fehlt / API-Fehler / Refusal / leer / Netzwerk -> Betreff-Titel + neutraler Hinweis,
  blockiert nie das Anhängen. Spec: `docs/superpowers/specs/2026-07-07-mail-zusammenfassung-design.md`.
  99 Tests. **GOTCHA: Datenschutz - der Mailinhalt (CMI-Geschäftsmail!) verlässt das Haus zu
  Anthropic. Bewusste Entscheidung Manuels (privater Key); das Add-in kontaktiert damit erstmals
  einen Dritten ausser Todoist. Bei CMI-Rollout/-Weitergabe Compliance beachten.** GOTCHA: der
  Anthropic-Key liegt (wie der Todoist-Token) im Browser/roamingSettings, technisch auslesbar,
  roamt übers Postfach. GOTCHA: Anthropic-Browser-Call braucht Header
  `anthropic-dangerous-direct-browser-access: true` (CORS-Preflight verifiziert `allow-origin: *`).


- 2026-07-06 spät (3): **ECHTE Ursache gefunden (Instrumentierung via Todoist-App-Log + PrintWindow).**
  Das Todoist-App-Log (`%APPDATA%/Todoist/logs/todoist-main.log`) protokolliert jeden
  empfangenen Deeplink ("Received deeplink: ..."). Damit verifiziert:
  (1) `todoist://task?id=<NEUE-alphanumerische-v1-Id>` oeffnet die KORREKTE Task (per
  PrintWindow-Screenshot bestaetigt: Task-Detail "Besprechung mit SSU" offen). Die neue
  Id ist also richtig; die alte numerische lehnt der Client ab (kein Foreground) -> der
  gestrige "35f3fa5"-Fix (id_mappings -> numerisch) war KOMPLETT FALSCH und wurde
  revertiert (getOldTaskIds + &web=-Param + taskDeepLink-2-Arg alles raus).
  (2) Der WAHRE Bug: die Redirect-Seite feuerte `todoist://` automatisch beim Load UND
  rief nach 1.5s `window.close()` auf. Chrome zeigt beim ersten externen Protokoll-Start
  pro Origin einen Bestaetigungsdialog; das Selbst-Schliessen killte den Dialog, bevor der
  Nutzer bestaetigen konnte -> App-Log blieb bei Button-Klicks LEER (kein Deeplink kam an).
  Beweis: Auto-Fire ohne window.close lieferte den Deeplink sofort; mit window.close/Live-
  Seite kam nichts. Fix: `open-task.html` ruft KEIN window.close mehr auf und zeigt einen
  prominenten Klick-Button "Task in Todoist öffnen" (Nutzergeste oeffnet zuverlaessig,
  Manuel bestaetigt: "mit Klick öffnet sich die Aufgabe") + Hinweis auf den Browser-Dialog.
  **GOTCHA 5: Protokoll-Redirect-Seiten NIE per window.close() selbst schliessen - das
  bricht Chromes Erlaubnis-Dialog ab. Auto-Fire versuchen + Klick-Button stehen lassen.**
  **GOTCHA 6: Zum Debuggen von todoist://-Deeplinks das App-Log tailen; es zeigt genau,
  welche URL ankam und ob sie "handled" oder "Cannot handle" wurde.**
- 2026-07-06 spät (4): **Deprecation-Warnung "Link aktualisieren" beseitigt. GELÖST + von Manuel on-device bestätigt ("jetzt funktioniert es").** `todoist://task?id=`
  oeffnet zwar die richtige Task, gilt aber als altes Link-Format (Todoists "Aktualisiere veraltete
  Aufgabenlinks", showTask?id=<v1_id>) und zeigt JEDES Mal die Warnung. Loesung aus dem App-Bundle
  (`resources/app.asar`, `deepLinkEndpoints` + `lt(pathname)`): `todoist://navigate-to?url=<neue-URL>`
  navigiert direkt ueber das NEUE Task-URL-Format `https://app.todoist.com/app/task/<id>` (die
  einzige gueltige Task-URL laut Doku) und umgeht den alten Resolver -> keine Warnung (Log:
  "Navigation deeplink successfully navigated to: task", Screenshot bestaetigt korrekte Task ohne
  Dialog). Host ist `navigate-to` (Bindestrich!), url muss same-origin app.todoist.com sein.
  **GOTCHA 7: Fuer todoist://-Deeplinks IMMER `navigate-to?url=https://app.todoist.com/app/task/<id>`
  statt `task?id=<id>` - letzteres triggert die "Link aktualisieren"-Deprecation-Warnung.
  Die Desktop-App registriert nur todoist:// (kein https-Universal-Link), Desktop nur so erreichbar.**
- 2026-07-06 spät (2, REVERTIERT): ~~Deep-Link-Fix (`35f3fa5`): todoist://task?id= braucht die ALTE numerische Id.~~ Falsch, siehe oben.
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
  **GOTCHA 4 (Ursache für Manuels erneuten Fehlversuch): `npm run deploy` meldete
  "Published", aber der GitHub-Pages-BUILD hing ~18 Min auf "building" (alle sonstigen
  Builds < 1 Min). Der gh-pages-Branch hatte den Fix (raw.githubusercontent lieferte ihn),
  aber der Pages-CDN (`manuelweingartner.github.io`) servierte weiter die ALTE Version.
  Manuel testete also alten Code. Verifikation NIE nur gegen den Branch, immer gegen die
  Pages-CDN-URL (`curl manuelweingartner.github.io/.../open-task.html`), und Build-Status
  via `gh api repos/.../pages/builds/latest`. Recovery: Build neu antriggern mit
  `gh api -X POST repos/manuelweingartner/outlook-todoist-addin/pages/builds`.**
  Offen/UNVERIFIZIERT: (1) Ob `todoist://task?id=<alte-numerische-id>` im Desktop-Client
  9.29.1 wirklich zur Task navigiert. On-Device bestätigt: `todoist://today` und
  `todoist://project?id=` navigieren, aber der Fenstertitel ist als Signal für die
  Task-Detail-Ansicht unzuverlässig. Der ALTE Live-Code feuerte bereits
  `todoist://task?id=<neue-alphanumerische-id>` und Manuel meldete "öffnet App, nicht Task"
  -> neue Id navigiert nicht. Ob die numerische Id via id_mappings das behebt, braucht
  einen Token-Test (Manuels Todoist-Token, secure stdin, nicht Chat).
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
