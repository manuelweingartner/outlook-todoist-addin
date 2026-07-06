# Design: Verbesserungspaket (Pfeiltasten, Neuer-Task-Formular, Ablage-Titel, #Projekt-Suche, Robustheit)

Datum 2026-07-06. Fünf mit Manuel abgestimmte Verbesserungen (aus der 10er-Vorschlagsliste:
Nr. 4, 6, 8, 9, 10). UX-Entscheidungen von Manuel: Fälligkeit = Schnell-Chips + Freitext;
25MB-Fall = expliziter Button "Ohne Anhänge anhängen"; Ablage-Titel = "Betreff (Datum, von Name)".

## 1. Pfeiltasten-Navigation (Vorschlag 4)

**Verhalten:** Fokus bleibt im Suchfeld. ↑/↓ bewegen eine sichtbare Markierung durch alle
gerenderten Task-Zeilen (über Sektionsgrenzen hinweg: Vorschläge, Überfällig, Heute bzw.
Treffer), mit `scrollIntoView({block: "nearest"})`. Enter hängt an die markierte Zeile an;
ohne explizite Markierung wie bisher an die oberste. Bei jedem Neu-Rendern (Tippen,
Load-Ende) springt die Markierung auf die oberste Zeile zurück.

**Umsetzung:**
- Reine Funktion in `taskLogic.ts`: `moveSelection(current: number, delta: number, count: number): number`
  (clampt auf [0, count-1]; count 0 liefert -1).
- `taskpane.ts`: Modul-State `selectedIndex`, Keydown-Handler am Suchfeld erweitert
  (ArrowUp/ArrowDown mit `preventDefault`), CSS-Klasse `.selected` auf der `task-row`
  (sichtbarer Hintergrund/Umrandung, Light + Dark). Render setzt `selectedIndex = 0`.
  Enter klickt die Zeile mit `.selected` (Fallback: erste).

## 2. Neuer-Task-Formular (Vorschlag 6)

**Verhalten:** Der Button "+ Neuen Task mit dieser Mail" klappt ein Inline-Formular auf
(zweiter Klick / Abbrechen klappt zu):
- Titel: Textfeld, vorbefüllt mit dem Mail-Betreff, editierbar.
- Projekt: Dropdown aus den bereits geladenen Projekten, Default Inbox
  (Projekt mit `is_inbox_project` bzw. erster Eintrag; ohne Projekte: kein Dropdown-Zwang,
  Feld weglassbar -> Todoist legt in Inbox an).
- Priorität: 4 wählbare Punkte P1-P4 in den bestehenden Prio-Farben, Default P4 (keine).
- Fälligkeit: Chips "Heute" / "Morgen" / "Nächste Woche" / "Kein Datum" (Default) plus
  Freitextfeld. Chips erzeugen deterministisch ein `due_date` (heute+0/+1/+7 via `todayIso`).
  Ist das Freitextfeld nicht leer, gewinnt es und geht als `due_string` mit `due_lang: "de"`
  an die API (Todoist parst natürliche Sprache).
- Aktionen: "Erstellen + Mail anhängen" (erstellt Task, hängt die vorbereitete Mail an,
  lädt die Liste neu, klappt zu) und "Abbrechen".

**Umsetzung:**
- `todoist.ts`: `createTask(token, options: NewTaskOptions): Promise<TodoistTask>` mit
  `NewTaskOptions = { content: string; project_id?: string; priority?: number;
  due_date?: string; due_string?: string; due_lang?: string }`. (Signaturwechsel; der
  bisherige Aufrufer wird mitgezogen.)
- Reine Funktion in `taskLogic.ts`:
  `buildNewTaskOptions(input: { title: string; projectId: string | null; priority: number;
  dueChip: "today" | "tomorrow" | "nextWeek" | "none"; dueText: string; today: string }): NewTaskOptions`
  (trimmt Titel, leerer Titel -> Fallback "Mail"; Freitext gewinnt über Chip; P4 wird als
  `priority: 1` gesendet bzw. weggelassen).
- `taskpane.html/css`: Formular-Markup + Stile (Todoist-Look, Dark-Mode-Variablen).

## 3. Ablage-Titel mit Absender (Vorschlag 8)

**Verhalten:** Kommentartext der abgelegten Mail neu `Betreff (Datum, von Name)`,
z.B. `Rechnung Q3 Hosting (05.07.2026, von Joel Willi)`. Name = Anzeigename des Absenders;
hat der Absender keinen Anzeigenamen, die Mail-Adresse; fehlt der Absender ganz, bleibt
das bisherige Format `Betreff (Datum)`. Fehlt auch das Datum: nur Betreff bzw.
`Betreff (von Name)`.

**Umsetzung:**
- Reine Funktion in `attachToTask.ts`: `senderName(from: string): string`, extrahiert aus
  dem bereits formatierten `MailData.from` ("Name <adresse>") den Namen; ist der String
  nur eine Adresse, wird sie zurückgegeben; leer -> "".
- `prepareCurrentMail`: Kommentartext-Bau erweitert (Datum und Name unabhängig optional).

## 4. #Projekt-Suchsyntax (Vorschlag 9)

**Verhalten:** Suchwörter mit führendem `#` matchen NUR den Projektnamen (case-insensitiv,
Substring); Wörter ohne `#` wie bisher Titel ODER Projektname; alle Wörter UND-verknüpft.
`#sap rechnung` = Task liegt in einem Projekt, dessen Name "sap" enthält, UND "rechnung"
kommt in Titel oder Projektname vor. Ein nacktes `#` (ohne Text dahinter) wird ignoriert.

**Umsetzung:** Erweiterung von `filterTasks` in `taskLogic.ts` (Signatur unverändert).

## 5. Robustheit (Vorschlag 10)

**(a) maxPages-Guard:** `getAllTasks` bricht nach `MAX_PAGES = 50` Seiten (10'000 Tasks)
ab, loggt via `console.error` und liefert das bis dahin Geladene zurück (funktional
degradieren statt endlos loopen, falls die API einen stabilen Cursor liefert).

**(b) Retry statt Token-Screen:** Scheitert `loadTasks`, wird unterschieden:
- `TodoistError` mit Status 401/403 -> wie heute Token-Section (Token ungültig).
- Alles andere (Netzfehler `Failed to fetch`, 5xx, Timeout) -> Fehlerbanner im Pane mit
  Button "Erneut versuchen" (ruft `loadTasks` erneut auf). Fehler wird weiterhin geloggt
  und im Status gezeigt (keine stillen Catches).

**(c) Text-only-.eml bei >25 MB:** Die Grössen-Warnung erhält den Button
"Ohne Anhänge anhängen". Klick baut die .eml ohne Datei-Anhänge neu (Header + Body bleiben),
ersetzt den vorbereiteten Mail-Zustand, aktiviert die Task-Zeilen und den Neuen-Task-Button
wieder und ersetzt die Warnung durch den Hinweis "Anhänge werden weggelassen.".
Ist auch die Text-only-Variante über 25 MB (Riesen-Body), bleibt alles deaktiviert wie heute.

**Umsetzung:**
- `attachToTask.ts`: neue exportierte reine Funktion
  `prepareMail(mail: MailData, includeAttachments: boolean): PreparedMail`;
  `prepareCurrentMail` bleibt als Convenience erhalten und ruft `prepareMail(mail, true)`.
  Die rohe `MailData` wird zusätzlich im Pane-State gehalten, damit der Button die
  Text-only-Variante ohne zweites Office.js-Lesen bauen kann (dafür liefert eine neue
  Funktion `readAndPrepareCurrentMail(): Promise<{ mail: MailData; prepared: PreparedMail }>`
  beides; `prepareCurrentMail` delegiert darauf).
- `taskpane.ts`: State `mailData: MailData | null`; Warnungs-Rendering mit Button;
  Hinweis-Text bei aktiver Text-only-Variante.

## 6. Desktop-Deeplinks (Nachtrag, Manuel 2026-07-06)

**Verhalten:** Alle Task-Links öffnen den installierten Todoist-Desktop-Client via
URL-Schema `todoist://task?id=<id>` statt app.todoist.com im Browser: die bestehenden
↗-Pfeile in jeder Zeile UND ein neuer Link "In Todoist öffnen" neben "Rückgängig" nach
erfolgreichem Anhängen. Vorbehalt: Ob die Outlook-Webview Custom-Protokolle durchlässt,
zeigt der on-device-Test; falls nicht, Revert auf Web-Links (Ein-Zeilen-Änderung).

**Umsetzung:** `taskDeepLink(id)` in `taskLogic.ts` liefert neu `todoist://task?id=<id>`
(Test anpassen). Navigation über Anchor-Click-Helfer `openExternal(url)` in `taskpane.ts`
(window.open ist bei Custom-Protokollen unzuverlässig). `renderUndo` erhält die Task-Id
und rendert den Zusatz-Link im Undo-Stil.

## Fehlerbehandlung

- Formular: Fehler beim Erstellen/Anhängen wie bisher via `setStatus(..., "err", e)`;
  Formular bleibt offen (Eingaben gehen nicht verloren).
- `due_string`, das Todoist nicht versteht: API-Fehler wird angezeigt (kein stilles Wegwerfen).
- `senderName` und `buildNewTaskOptions` sind total (liefern für jeden Input ein Ergebnis).

## Tests (TDD)

- `moveSelection`: clamp oben/unten, count 0, normale Schritte.
- `buildNewTaskOptions`: Chip-Mapping (+0/+1/+7), Freitext gewinnt, P4 weggelassen,
  leerer Titel -> "Mail", Projekt optional.
- `senderName`: "Name <adr>" -> Name, nur Adresse -> Adresse, leer -> "".
- Kommentartext: alle vier Kombinationen (Datum/Name je da/fehlt).
- `filterTasks`: #-Token matcht nur Projekt, Kombination #+Wort, nacktes #, Case.
- `getAllTasks`: bricht nach MAX_PAGES ab und liefert Teilmenge (fetch-Mock mit stabilem Cursor).
- `prepareMail`: includeAttachments false -> .eml ohne Attachment-Parts, Grösse kleiner,
  Header/Body erhalten.
- Retry-Weiche: 401 -> Token-Pfad, sonst Retry-Pfad (reine Klassifizier-Funktion
  `isAuthError(e): boolean` in `taskLogic.ts` oder `todoist.ts`, getestet).

## Nicht-Ziele

- Kein Datums-Picker, keine Labels/Sections im Formular (YAGNI).
- Keine Retry-Automatik mit Backoff (nur manueller Button).
- Keine Kompression/Aufteilung übergrosser Mails.
- `@label`-Syntax in der Suche kommt NICHT (nur `#projekt`).
