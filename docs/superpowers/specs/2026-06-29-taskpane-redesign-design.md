# Design: Task-Pane-Redesign „Mail an Todoist"

Datum: 2026-06-29
Status: Genehmigt (Brainstorming abgeschlossen)
Autor: Manuel Weingartner

## Problem

Das Add-in funktioniert (zentral ausgerollt, Ticket #140204), die Seitenleiste ist
aber noch das nackte Microsoft-Starter-Template: Inline-Styles, ungenutztes
`.ms-welcome`-Skelett, kein echtes Design, keine sauberen Zwischenzustaende. Es soll
optisch und in der Bedienung „schoen" werden, im Todoist-Look.

### Harte Rahmenbedingung: kein Manifest-Eingriff

Die IT/Ops hat nur das `manifest.prod.xml` ausgerollt (Verweis auf die GitHub-Pages-URL
+ Ribbon-Button). Das UI (HTML/CSS/JS) liegt live auf GitHub Pages und wird bei jedem
Oeffnen frisch geladen. **Alle Aenderungen dieser Spec betreffen nur das gehostete UI
und werden per `npm run deploy` allein vom Autor ausgerollt.** Das Manifest (Button-Label,
Icon-URLs, Permissions, `RequestedHeight` 250px) wird NICHT angefasst, damit kein
erneuter IT-Rollout noetig ist.

## Visuelles System (Todoist-Branding)

- Akzentfarbe Todoist-Rot `#E44332`, neutrale Graustufen drumherum.
- System-Font-Stack (Segoe UI zuerst, in Outlook nativ).
- **Fluent-UI-CDN entfernen** (`fabric.min.css`): wird aktuell geladen, aber kaum genutzt.
  Self-contained CSS = schneller, eine externe Abhaengigkeit weniger, volle Kontrolle.
- **Dark-Mode** via `prefers-color-scheme: dark` (Outlook hat einen Dunkelmodus, das Pane
  geht mit). Farben als CSS-Custom-Properties, je ein Light-/Dark-Set.
- Konventionen aus CLAUDE.md gelten: keine Em-/En-Dashes, echte Umlaute, kein stilles
  `try/catch` (Fehler immer `console.error` UND sichtbar im Pane).

## Die fuenf Zustaende (UX-Feinschliff)

1. **Onboarding / Token:** zentrierter Screen mit kurzer Erklaerung, dem genauen
   Todoist-Pfad als Schritt-Hinweis, Passwort-Feld, rotem „Speichern"-Button. Ersetzt den
   heutigen nackten Text + Input.
2. **Laden:** Skeleton-Platzhalter-Zeilen (graue Balken) statt leerer Liste/Spinner.
3. **Task-Liste:** echte Liste mit Hover, gruppiert in „Heute" / „Ueberfaellig",
   Prioritaets-Punkt links, Projektname dezent rechts, Suchfeld mit Lupe oben.
4. **Leer:** freundlicher Leerzustand („Keine Tasks fuer heute") statt Status-Textzeile.
5. **Anhaengen:** pro Zeile Inline-Feedback (Spinner beim Klick -> gruener Haken
   „angehaengt"), Fehler weiter rot inline.

## Interaktion

**Ein-Klick-Anhaengen bleibt** (schnellster Weg, Daseinsberechtigung des Add-ins). Jede
Zeile ist ein klar klickbarer Button mit Hover; das sofortige Inline-Feedback +
Rueckgaengig (s.u.) nimmt dem Fehlklick das Risiko. Kein separater Bestaetigen-Button.

## Zusatz-Features (genehmigt: alle ausser „Token aendern/abmelden")

1. **Rueckgaengig nach dem Anhaengen**: nach Erfolg ein „Rueckgaengig"-Link, der den
   gerade erstellten Kommentar wieder loescht (`DELETE /comments/{id}`).
   Erfordert: `addComment` gibt die neue Kommentar-`id` zurueck (heute `void`);
   neue `deleteComment(token, id)` in `todoist.ts`.
2. **Betreff statt Dateiname als Kommentartext**: Kommentartext wird „<Betreff>,
   <Datum>", nicht der `.eml`-Dateiname. Betreff/Datum
   kommen aus `mailReader` (Office.js `item.subject`, `item.dateTimeCreated`).
3. **Mail-Kontext-Kopf**: fixe Kopfzeile „Anhaengen: ‚<Betreff>'", zeigt welche Mail
   gerade angehaengt wird. Quelle wie (2).
4. **Vorab-Groessencheck (25 MB)**: der bestehende 25-MB-Check (in `attachToTask`) wird
   vorgezogen: nach `.eml`-Bau die Groesse ermitteln und, falls zu gross, dezent oben
   warnen und die Task-Klicks deaktivieren, statt erst beim Anhaengen zu scheitern.
5. **Loading-Skeleton**: siehe Zustand 2 (graue Platzhalter-Zeilen).
6. **Tastatur-Flow**: Suchfeld bekommt beim Oeffnen Fokus; `Enter` haengt an den obersten
   Treffer an. Liste per Pfeiltasten fokussierbar (native Button-Fokus-Reihenfolge reicht).
7. **Projektname statt project_id**: einmaliger `/projects`-Abruf, Map `id -> name`,
   Projektname als dezentes Label pro Task. Neue `getProjects(token)` in `todoist.ts`,
   `TodoistTask` erhaelt optional `priority` und `due` (s.u.).
8. **„In Todoist oeffnen"-Link pro Task**: kleiner Pfeil/Link je Task, oeffnet den Task in
   Todoist (`https://app.todoist.com/app/task/<id>` in neuem Tab via `window.open`).
9. **Neuen Task aus der Mail erstellen**: falls kein passender Task: „+ Neuen Task mit
   dieser Mail". `POST /tasks` mit Betreff als content, danach gleicher Anhaenge-Flow.
   Neue `createTask(token, content)` in `todoist.ts`.

## Daten-Aenderungen (klein, API liefert die Felder bereits)

`TodoistTask` wird erweitert: `priority` (1 bis 4, 4 = hoechste) und `due` (Objekt mit
`date`, optional `datetime`). Beide kommen aus der bestehenden v1-Antwort von
`/tasks/filter`, werden heute nur nicht gemappt. Daraus:
- Gruppierung „Heute" / „Ueberfaellig" (Vergleich `due.date` mit heute).
- Prioritaets-Punkt-Farbe (p1 rot, p2 orange, p3 blau, p4 grau; Todoist-Konvention,
  wobei API-`priority` 4 = p1).

## Betroffene Module

- `src/taskpane/taskpane.html`: Markup fuer die fuenf Zustaende, Kontext-Kopf,
  Fluent-Link raus, semantische Struktur statt Inline-Styles.
- `src/taskpane/taskpane.css`: komplettes Neuschreiben: Todoist-Theme, CSS-Variablen,
  Light/Dark, Skeleton, Hover/Fokus, Prioritaets-Punkte.
- `src/taskpane/taskpane.ts`: Zustands-Rendering, Gruppierung, Inline-Anhaenge-Feedback,
  Rueckgaengig, Tastatur-Flow, Groessen-Vorabcheck, Projektnamen-Map, „neuer Task".
- `src/lib/todoist.ts`: `TodoistTask` um `priority`/`due` erweitern; `addComment` gibt
  `id` zurueck; neue `deleteComment`, `getProjects`, `createTask`.
- `src/lib/mailReader.ts`: Betreff + Datum mit ausliefern (fuer Kommentartext + Kopf).
- `src/lib/attachToTask.ts`: Groessencheck so umbauen, dass die Groesse vorab abfragbar
  ist; Kommentartext aus Betreff/Datum; Rueckgabe der Kommentar-id durchreichen.

## Tests (TDD, Pflicht laut CLAUDE.md)

- `todoist.ts`: `deleteComment`/`getProjects`/`createTask` (fetch-Mock, CORS-/Error-Pfade),
  `addComment` gibt id zurueck, Mapping von `priority`/`due`.
- `attachToTask.ts`: Groessencheck-Vorabpfad (zu gross -> Fehler vor Upload), Kommentartext
  = Betreff+Datum, id-Durchreichung, Undo-Pfad ruft `deleteComment`.
- `mailReader.ts`: Betreff/Datum-Extraktion.
- Reine UI-Optik (CSS/Dark-Mode) wird nicht unit-getestet; Logik in `taskpane.ts`
  (Gruppierung „heute/ueberfaellig", Prioritaets-Mapping) als reine Funktionen ausgelagert
  und getestet.

## Nicht-Ziele / bewusst weggelassen

- **Kein „Token aendern/abmelden"** (Vorschlag 4 ausgelassen).
- Kein Manifest-Eingriff, keine neuen Permissions, kein Microsoft Graph.
- Kein Mehrsprachen-UI (nur Deutsch).
- Keine Aenderung am `.eml`-Bau selbst.

## Risiken / offene Punkte

- `due`/`priority`-Feldnamen der v1-API empirisch verifizieren (gegen echte Antwort),
  bevor das Mapping als fix gilt (Konvention „konkrete Fakten verifizieren").
- Todoist-Task-Deeplink-URL (`app.todoist.com/app/task/<id>`) am echten Account testen.
- Skeleton-Hoehe vs. `RequestedHeight` 250px im Manifest: das Pane ist scrollbar, kein
  Manifest-Eingriff noetig; nur sicherstellen, dass nichts ueber 250px „fest" sein muss.
