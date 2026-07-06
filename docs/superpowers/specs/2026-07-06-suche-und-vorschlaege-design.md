# Design: Bessere Suche + Top-3-Vorschläge (2026-07-06)

## Problem

1. Die manuelle Task-Suche läuft server-seitig über Todoists Filter-Query
   (`/tasks/filter?query=search: <text>`) und matcht case-sensitiv: "sap" findet
   "SAP" nicht. Das ist für den Hauptflow (Task finden, Mail anhängen) ein harter
   Usability-Bruch.
2. Der Nutzer muss den passenden Task immer selbst suchen. Das Add-in kennt aber
   Betreff und Inhalt der offenen Mail und könnte passende Tasks direkt vorschlagen.
3. Nebenbefund: `unwrap()` liest nur die erste Seite der paginierten v1-Antwort.
   `next_cursor` wird ignoriert, auch die heutige Heute/Überfällig-Liste ist also
   bei vielen Tasks unvollständig.

## Entscheidung (mit Manuel abgestimmt)

- **Ansatz A: alles client-seitig.** Beim Pane-Start werden einmal ALLE offenen
  Tasks geladen (paginiert). Suche und Vorschläge laufen danach rein lokal.
  Verworfen: Server-Suche normalisieren (Matching nicht kontrollierbar, halbe
  Lösung) und Hybrid (Komplexität ohne Nutzen).
- **Vorschläge basieren auf Betreff + Mail-Inhalt**, Betreff stärker gewichtet.
- **Suchraum = alle offenen Tasks**, nicht nur Heute/Überfällig.

## Architektur

### 1. Datenbeschaffung: `getAllTasks` (src/lib/todoist.ts)

- Neues `getAllTasks(token): Promise<TodoistTask[]>`: GET `/tasks?limit=200`,
  folgt `next_cursor` bis `null`, konkateniert alle Seiten.
- Wird beim Pane-Start EINMAL geladen und im Modul-State des Panes gehalten.
- `searchTasks` (server-seitig) wird nicht mehr vom Pane benutzt und entfernt.
- Die Heute/Überfällig-Standardansicht wird aus dem Voll-Datensatz client-seitig
  gefiltert (bestehendes `groupTasks` + `todayIso`, Verhalten wie heute).

### 2. Suche: `filterTasks` (src/taskpane/taskLogic.ts, reine Logik)

`filterTasks(tasks, query, projectNames): TodoistTask[]`

- Case-insensitiv (`toLowerCase()` beidseitig).
- Query wird an Whitespace gesplittet, mehrere Wörter = UND-Verknüpfung.
- Jedes Wort muss Substring von Task-Titel ODER Projektname sein
  ("sap" findet auch Tasks im Projekt "SAP").
- Leere Query = alle Tasks (Aufrufer zeigt dann die Standardansicht).
- Läuft direkt beim Tippen ohne API-Call; das 300ms-Debounce entfällt.
- Enter hängt weiterhin den obersten Treffer an (bestehendes Verhalten).

### 3. Vorschläge: `extractMailKeywords` + `suggestTasks` (taskLogic.ts, rein)

`extractMailKeywords(subject, bodyText): { subjectWords, bodyWords }`

- Betreff: `RE:/AW:/FW:/WG:`-Präfixe (auch gestapelt) entfernen.
- Body: auf die ersten 2000 Zeichen gekürzt (Aufrufer liefert bereits Text).
- Tokenisierung an Nicht-Wort-Grenzen, alles lowercase.
- Stoppwörter (kleine DE/EN-Liste: der/die/das/und/für/the/for/and/...) und
  Wörter unter 3 Zeichen fliegen raus. Duplikate dedupliziert; Wörter, die schon
  im Betreff sind, zählen nicht nochmal als Body-Wort.

`suggestTasks(tasks, keywords): TodoistTask[]`

- Score pro Task (case-insensitives Substring-Match gegen Task-Titel):
  Betreff-Wort = 3 Punkte, Body-Wort = 1 Punkt.
- Top 3 nach Score absteigend, Score 0 wird nie vorgeschlagen.
- Deterministische Reihenfolge bei Gleichstand (stabile Sortierung).

### 4. Mail-Body als Text: `PreparedMail.bodyText` (src/lib/attachToTask.ts)

- `MailData.htmlBody` wird zu Klartext gestrippt (Tags raus, Entities dekodiert,
  Whitespace kollabiert) und als neues Feld `bodyText` an `PreparedMail` gehängt.
- Strip-Funktion als reine, testbare Funktion (kein DOM nötig, Regex-basiert
  reicht für Scoring-Zwecke).

### 5. UI (src/taskpane/taskpane.{ts,html,css})

- Neue Sektion **"Vorschläge"** ZUOBERST (über Überfällig/Heute), gleiche
  Zeilen-Optik (`makeRow` wiederverwendet).
- Erscheint nur, wenn (a) das Suchfeld leer ist und (b) mindestens ein Vorschlag
  Score > 0 hat. Sobald getippt wird, ersetzt die Trefferliste die komplette
  Ansicht (Vorschläge weg).
- Trefferansicht bei aktiver Suche: flache Liste unter Label "Treffer"
  (keine Heute/Überfällig-Gruppierung, da über alle Tasks gesucht wird).

## Fehlerbehandlung

- Voll-Load scheitert: bestehender Fehlerpfad (Meldung + Token-Section), wie heute.
- Body nicht lesbar / leer: Vorschläge basieren nur auf dem Betreff; nie ein
  harter Fehler. Kein stilles Catch: Fehler wird geloggt (Repo-Konvention).
- 0 Vorschläge: Sektion wird schlicht nicht gerendert.

## Tests (TDD, Repo-Pflicht)

- `filterTasks`: Case-Insensitivität ("sap" findet "SAP"), Mehrwort-UND,
  Projektnamen-Match, leere Query.
- `extractMailKeywords`: Präfix-Bereinigung (auch "AW: RE:"), Stoppwörter,
  Mindestlänge, Dedupe, 2000-Zeichen-Cut.
- `suggestTasks`: Gewichtung 3/1, Top-3-Cut, Score-0-Ausschluss, stabile
  Reihenfolge.
- HTML-Strip: Tags, Entities, Whitespace.
- `getAllTasks`: Cursor-Loop über mehrere Seiten (fetch gemockt), Einzelseite.

## Nicht-Ziele

- Kein Fuzzy-Matching / keine Tippfehler-Toleranz (YAGNI, Substring reicht).
- Keine Suche in Task-Beschreibungen oder Kommentaren (v1-Task-Payload hat nur
  `content` geladen).
- Kein Caching über Pane-Sessions hinweg (Load pro Öffnen ist schnell genug).
