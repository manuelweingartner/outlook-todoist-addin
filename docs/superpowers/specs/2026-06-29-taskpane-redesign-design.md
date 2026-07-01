# Design: Task-Pane-Redesign „Mail an Todoist"

Datum: 2026-06-29
Status: Genehmigt (Brainstorming abgeschlossen)
Autor: Manuel Weingartner

## Problem

Das Add-in funktioniert (zentral ausgerollt, Ticket #140204), die Seitenleiste ist
aber noch das nackte Microsoft-Starter-Template: Inline-Styles, ungenutztes
`.ms-welcome`-Skelett, kein echtes Design, keine sauberen Zwischenzustände. Es soll
optisch und in der Bedienung „schön" werden, im Todoist-Look.

### Harte Rahmenbedingung: kein Manifest-Eingriff

Die IT/Ops hat nur das `manifest.prod.xml` ausgerollt (Verweis auf die GitHub-Pages-URL
+ Ribbon-Button). Das UI (HTML/CSS/JS) liegt live auf GitHub Pages und wird bei jedem
Öffnen frisch geladen. **Alle Änderungen dieser Spec betreffen nur das gehostete UI
und werden per `npm run deploy` allein vom Autor ausgerollt.** Das Manifest (Button-Label,
Icon-URLs, Permissions, `RequestedHeight` 250px) wird NICHT angefasst, damit kein
erneuter IT-Rollout nötig ist.

## Visuelles System (Todoist-Branding)

- Akzentfarbe Todoist-Rot `#E44332`, neutrale Graustufen drumherum.
- System-Font-Stack (Segoe UI zuerst, in Outlook nativ).
- **Fluent-UI-CDN entfernen** (`fabric.min.css`): wird aktuell geladen, aber kaum genutzt.
  Self-contained CSS = schneller, eine externe Abhängigkeit weniger, volle Kontrolle.
- **Dark-Mode** via `prefers-color-scheme: dark` (Outlook hat einen Dunkelmodus, das Pane
  geht mit). Farben als CSS-Custom-Properties, je ein Light-/Dark-Set.
- Konventionen aus CLAUDE.md gelten: keine Em-/En-Dashes, echte Umlaute, kein stilles
  `try/catch` (Fehler immer `console.error` UND sichtbar im Pane).

## Die fünf Zustände (UX-Feinschliff)

1. **Onboarding / Token:** zentrierter Screen mit kurzer Erklärung, dem genauen
   Todoist-Pfad als Schritt-Hinweis, Passwort-Feld, rotem „Speichern"-Button. Ersetzt den
   heutigen nackten Text + Input.
2. **Laden:** Skeleton-Platzhalter-Zeilen (graue Balken) statt leerer Liste/Spinner.
3. **Task-Liste:** echte Liste mit Hover, gruppiert in „Heute" / „Überfällig",
   Prioritäts-Punkt links, Projektname dezent rechts, Suchfeld mit Lupe oben.
4. **Leer:** freundlicher Leerzustand („Keine Tasks für heute") statt Status-Textzeile.
5. **Anhängen:** pro Zeile Inline-Feedback (Spinner beim Klick -> grüner Haken
   „angehängt"), Fehler weiter rot inline.

## Interaktion

**Ein-Klick-Anhängen bleibt** (schnellster Weg, Daseinsberechtigung des Add-ins). Jede
Zeile ist ein klar klickbarer Button mit Hover; das sofortige Inline-Feedback +
Rückgängig (s.u.) nimmt dem Fehlklick das Risiko. Kein separater Bestätigen-Button.

## Zusatz-Features (genehmigt: alle ausser „Token ändern/abmelden")

1. **Rückgängig nach dem Anhängen**: nach Erfolg ein „Rückgängig"-Link, der den
   gerade erstellten Kommentar wieder löscht (`DELETE /comments/{id}`).
   Erfordert: `addComment` gibt die neue Kommentar-`id` zurück (heute `void`);
   neue `deleteComment(token, id)` in `todoist.ts`.
2. **Betreff statt Dateiname als Kommentartext**: Kommentartext wird „<Betreff>,
   <Datum>", nicht der `.eml`-Dateiname. Betreff/Datum
   kommen aus `mailReader` (Office.js `item.subject`, `item.dateTimeCreated`).
3. **Mail-Kontext-Kopf**: fixe Kopfzeile „Anhängen: ‚<Betreff>'", zeigt welche Mail
   gerade angehängt wird. Quelle wie (2).
4. **Vorab-Grössencheck (25 MB)**: der bestehende 25-MB-Check (in `attachToTask`) wird
   vorgezogen: nach `.eml`-Bau die Grösse ermitteln und, falls zu gross, dezent oben
   warnen und die Task-Klicks deaktivieren, statt erst beim Anhängen zu scheitern.
5. **Loading-Skeleton**: siehe Zustand 2 (graue Platzhalter-Zeilen).
6. **Tastatur-Flow**: Suchfeld bekommt beim Öffnen Fokus; `Enter` hängt an den obersten
   Treffer an. Liste per Pfeiltasten fokussierbar (native Button-Fokus-Reihenfolge reicht).
7. **Projektname statt project_id**: einmaliger `/projects`-Abruf, Map `id -> name`,
   Projektname als dezentes Label pro Task. Neue `getProjects(token)` in `todoist.ts`,
   `TodoistTask` erhält optional `priority` und `due` (s.u.).
8. **„In Todoist öffnen"-Link pro Task**: kleiner Pfeil/Link je Task, öffnet den Task in
   Todoist (`https://app.todoist.com/app/task/<id>` in neuem Tab via `window.open`).
9. **Neuen Task aus der Mail erstellen**: falls kein passender Task: „+ Neuen Task mit
   dieser Mail". `POST /tasks` mit Betreff als content, danach gleicher Anhänge-Flow.
   Neue `createTask(token, content)` in `todoist.ts`.

## Daten-Änderungen (klein, API liefert die Felder bereits)

`TodoistTask` wird erweitert: `priority` (1 bis 4, 4 = höchste) und `due` (Objekt mit
`date`, optional `datetime`). Beide kommen aus der bestehenden v1-Antwort von
`/tasks/filter`, werden heute nur nicht gemappt. Daraus:
- Gruppierung „Heute" / „Überfällig" (Vergleich `due.date` mit heute).
- Prioritäts-Punkt-Farbe (p1 rot, p2 orange, p3 blau, p4 grau; Todoist-Konvention,
  wobei API-`priority` 4 = p1).

## Betroffene Module

- `src/taskpane/taskpane.html`: Markup für die fünf Zustände, Kontext-Kopf,
  Fluent-Link raus, semantische Struktur statt Inline-Styles.
- `src/taskpane/taskpane.css`: komplettes Neuschreiben: Todoist-Theme, CSS-Variablen,
  Light/Dark, Skeleton, Hover/Fokus, Prioritäts-Punkte.
- `src/taskpane/taskpane.ts`: Zustands-Rendering, Gruppierung, Inline-Anhänge-Feedback,
  Rückgängig, Tastatur-Flow, Grössen-Vorabcheck, Projektnamen-Map, „neuer Task".
- `src/lib/todoist.ts`: `TodoistTask` um `priority`/`due` erweitern; `addComment` gibt
  `id` zurück; neue `deleteComment`, `getProjects`, `createTask`.
- `src/lib/mailReader.ts`: Betreff + Datum mit ausliefern (für Kommentartext + Kopf).
- `src/lib/attachToTask.ts`: Grössencheck so umbauen, dass die Grösse vorab abfragbar
  ist; Kommentartext aus Betreff/Datum; Rückgabe der Kommentar-id durchreichen.

## Tests (TDD, Pflicht laut CLAUDE.md)

- `todoist.ts`: `deleteComment`/`getProjects`/`createTask` (fetch-Mock, CORS-/Error-Pfade),
  `addComment` gibt id zurück, Mapping von `priority`/`due`.
- `attachToTask.ts`: Grössencheck-Vorabpfad (zu gross -> Fehler vor Upload), Kommentartext
  = Betreff+Datum, id-Durchreichung, Undo-Pfad ruft `deleteComment`.
- `mailReader.ts`: Betreff/Datum-Extraktion.
- Reine UI-Optik (CSS/Dark-Mode) wird nicht unit-getestet; Logik in `taskpane.ts`
  (Gruppierung „heute/überfällig", Prioritäts-Mapping) als reine Funktionen ausgelagert
  und getestet.

## Nicht-Ziele / bewusst weggelassen

- **Kein „Token ändern/abmelden"** (Vorschlag 4 ausgelassen).
- Kein Manifest-Eingriff, keine neuen Permissions, kein Microsoft Graph.
- Kein Mehrsprachen-UI (nur Deutsch).
- Keine Änderung am `.eml`-Bau selbst.

## Risiken / offene Punkte

- `due`/`priority`-Feldnamen der v1-API empirisch verifizieren (gegen echte Antwort),
  bevor das Mapping als fix gilt (Konvention „konkrete Fakten verifizieren").
- Todoist-Task-Deeplink-URL (`app.todoist.com/app/task/<id>`) am echten Account testen.
- Skeleton-Höhe vs. `RequestedHeight` 250px im Manifest: das Pane ist scrollbar, kein
  Manifest-Eingriff nötig; nur sicherstellen, dass nichts über 250px „fest" sein muss.
