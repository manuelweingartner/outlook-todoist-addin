# Verbesserungspaket Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Pfeiltasten-Navigation, Neuer-Task-Formular (Projekt/Priorität/Fälligkeit), Ablage-Titel mit Absender, #Projekt-Suchsyntax und Robustheit (maxPages, Retry-Banner, Text-only-.eml bei >25 MB).

**Architecture:** Alle neue Logik als reine, getestete Funktionen in `taskLogic.ts` / `todoist.ts` / `attachToTask.ts`; die UI-Verdrahtung in `taskpane.{ts,html,css}` kommt in drei getrennten Tasks am Schluss. `createTask` wechselt auf einen Options-Parameter (Aufrufer wird im selben Task mitgezogen, Build bleibt grün).

**Tech Stack:** TypeScript, Office.js, Todoist Unified API v1, Jest (ts-jest, jsdom), Webpack.

**Spec:** `docs/superpowers/specs/2026-07-06-verbesserungspaket-design.md`

## Global Constraints

- Git-Identität: Manuel Weingartner <manuel.weingartner@gmx.ch>. NIE Co-Authored-By.
- Keine Em-/En-Dashes in Code, Kommentaren, Docs.
- Echte Umlaute (ä/ö/ü) in sichtbarem Deutsch (UI-Strings, Markdown-Prosa); Code-Kommentare/Testnamen folgen dem Repo-Muster (ae/ue erlaubt). Schweizer Deutsch, kein ß.
- Kein stilles try/catch: Fehler loggen UND dem Nutzer zeigen (setStatus-Muster).
- TDD: Test zuerst, laufen lassen (muss failen), dann Implementation.
- Kein `git add -A`: Dateien immer explizit stagen.
- Tests: `npx jest`, Build: `npm run build`.
- UI-Strings exakt: "Erneut versuchen", "Ohne Anhänge anhängen", "Anhänge werden weggelassen.", "Erstellen + Mail anhängen", "Abbrechen", "Kein Datum", "Heute", "Morgen", "Nächste Woche".

---

### Task 1: taskLogic: #Projekt-Filter + moveSelection

**Files:**
- Modify: `src/taskpane/taskLogic.ts` (Funktion `filterTasks` ersetzen, `moveSelection` anhängen)
- Test: `tests/taskLogic.test.ts`

**Interfaces:**
- Consumes: bestehendes `filterTasks(tasks, query, projectNames)` und Test-Helper `named(id, content, project_id?)` in `tests/taskLogic.test.ts`.
- Produces:
  - `filterTasks` (Signatur unverändert): Wörter mit führendem `#` matchen NUR den Projektnamen (Substring, case-insensitiv); nacktes `#` wird ignoriert; Wörter ohne `#` wie bisher Titel ODER Projektname; alles UND-verknüpft.
  - `moveSelection(current: number, delta: number, count: number): number` (clampt auf [0, count-1]; count <= 0 liefert -1). Task 5 benutzt exakt diese Signatur.

- [ ] **Step 1: Failing Tests schreiben**

In `tests/taskLogic.test.ts` (Import um `moveSelection` erweitern; Fixture `projects = { p1: "Inbox", p2: "SAP" }` existiert im `filterTasks`-describe):

```ts
test("#-Wort matcht nur den Projektnamen, nicht den Titel", () => {
  const tasks = [named("imTitel", "sap Lizenzen klären", "p1"), named("imProjekt", "Rechnung zahlen", "p2")];
  expect(filterTasks(tasks, "#sap", projects).map((x) => x.id)).toEqual(["imProjekt"]);
});

test("#Projekt kombiniert mit normalem Wort (UND)", () => {
  const tasks = [named("a", "Rechnung pruefen", "p2"), named("b", "Schulung planen", "p2"), named("c", "Rechnung zahlen", "p1")];
  expect(filterTasks(tasks, "#sap rechnung", projects).map((x) => x.id)).toEqual(["a"]);
});

test("nacktes # wird ignoriert", () => {
  const tasks = [named("a", "x"), named("b", "y")];
  expect(filterTasks(tasks, "#", projects)).toHaveLength(2);
});

test("#-Match ist case-insensitiv", () => {
  const tasks = [named("a", "Rechnung", "p2")];
  expect(filterTasks(tasks, "#SAP", projects).map((x) => x.id)).toEqual(["a"]);
});

describe("moveSelection", () => {
  test("normaler Schritt runter und rauf", () => {
    expect(moveSelection(1, 1, 5)).toBe(2);
    expect(moveSelection(2, -1, 5)).toBe(1);
  });
  test("clampt oben und unten", () => {
    expect(moveSelection(0, -1, 5)).toBe(0);
    expect(moveSelection(4, 1, 5)).toBe(4);
  });
  test("leere Liste liefert -1", () => {
    expect(moveSelection(0, 1, 0)).toBe(-1);
  });
  test("negativer Startindex landet bei 0", () => {
    expect(moveSelection(-1, 1, 3)).toBe(0);
  });
});
```

- [ ] **Step 2: Tests laufen lassen, muessen failen**

Run: `npx jest tests/taskLogic.test.ts`
Expected: FAIL (#-Tests matchen falsch, `moveSelection` fehlt).

- [ ] **Step 3: Implementieren**

`filterTasks` in `src/taskpane/taskLogic.ts` ersetzen durch:

```ts
// Client-seitige Suche: case-insensitiv, alle Woerter muessen matchen (UND).
// Woerter mit fuehrendem # matchen NUR den Projektnamen ("#sap rechnung"),
// Woerter ohne # matchen Task-Titel ODER Projektname. Nacktes # wird ignoriert.
export function filterTasks(
  tasks: TodoistTask[],
  query: string,
  projectNames: Record<string, string>,
): TodoistTask[] {
  const words = query.toLowerCase().split(/\s+/).filter(Boolean);
  if (words.length === 0) return tasks;
  return tasks.filter((task) => {
    const project = (projectNames[task.project_id] ?? "").toLowerCase();
    const hay = `${task.content.toLowerCase()} ${project}`;
    return words.every((w) => {
      if (w.startsWith("#")) {
        const p = w.slice(1);
        return p === "" ? true : project.includes(p);
      }
      return hay.includes(w);
    });
  });
}
```

`moveSelection` anhängen:

```ts
// Pfeiltasten-Auswahl: clampt auf [0, count-1]; ohne Zeilen gibt es keine Auswahl (-1).
export function moveSelection(current: number, delta: number, count: number): number {
  if (count <= 0) return -1;
  return Math.min(Math.max(current + delta, 0), count - 1);
}
```

- [ ] **Step 4: Tests laufen lassen, muessen passen**

Run: `npx jest tests/taskLogic.test.ts`
Expected: PASS (inkl. aller bestehenden filterTasks-Tests).

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/taskLogic.ts tests/taskLogic.test.ts
git commit -m "TaskLogic: #Projekt-Suchsyntax + moveSelection fuer Pfeiltasten"
```

---

### Task 2: todoist: createTask mit Options, maxPages-Guard, isAuthError

**Files:**
- Modify: `src/lib/todoist.ts` (createTask ersetzen, getAllTasks erweitern, isAuthError + NewTaskOptions neu)
- Modify: `src/taskpane/taskpane.ts` (einzige Zeile: createTask-Aufruf auf Options-Form)
- Test: `tests/todoist.test.ts`

**Interfaces:**
- Consumes: bestehende Helfer `auth`, `ensureOk`, `unwrap`, `API`, `TodoistError`, `Paginated<T>`.
- Produces (Tasks 3, 5, 6 verlassen sich exakt darauf):
  - `interface NewTaskOptions { content: string; project_id?: string; priority?: number; due_date?: string; due_string?: string; due_lang?: string }` (exportiert)
  - `createTask(token: string, options: NewTaskOptions): Promise<TodoistTask>` (POSTet options als JSON-Body)
  - `isAuthError(e: unknown): boolean` (true nur fuer TodoistError mit Status 401 oder 403)
  - `getAllTasks` bricht nach `MAX_PAGES = 50` Seiten ab (console.error + bis dahin Geladenes zurueckgeben)

- [ ] **Step 1: Failing Tests schreiben**

In `tests/todoist.test.ts` (Imports um `isAuthError` erweitern; `createTask` ist schon importiert). Bestehenden `createTask`-Test auf die Options-Form umstellen und erweitern:

```ts
describe("createTask", () => {
  test("POSTet Options als JSON-Body", async () => {
    mockFetch(200, { id: "9", content: "Neu", project_id: "p1" });
    const task = await createTask("tok", { content: "Neu", project_id: "p1", priority: 3, due_date: "2026-07-07" });
    expect(task.id).toBe("9");
    const [url, opts] = (global as any).fetch.mock.calls[0];
    expect(url).toBe("https://api.todoist.com/api/v1/tasks");
    expect(opts.method).toBe("POST");
    expect(JSON.parse(opts.body)).toEqual({ content: "Neu", project_id: "p1", priority: 3, due_date: "2026-07-07" });
  });
  test("minimaler Aufruf nur mit content", async () => {
    mockFetch(200, { id: "9", content: "Neu", project_id: "p1" });
    await createTask("tok", { content: "Neu" });
    expect(JSON.parse((global as any).fetch.mock.calls[0][1].body)).toEqual({ content: "Neu" });
  });
});

describe("isAuthError", () => {
  test("401 und 403 sind Auth-Fehler", () => {
    expect(isAuthError(new TodoistError(401, "unauthorized"))).toBe(true);
    expect(isAuthError(new TodoistError(403, "forbidden"))).toBe(true);
  });
  test("andere Fehler nicht", () => {
    expect(isAuthError(new TodoistError(500, "boom"))).toBe(false);
    expect(isAuthError(new TypeError("Failed to fetch"))).toBe(false);
    expect(isAuthError(undefined)).toBe(false);
  });
});

describe("getAllTasks maxPages", () => {
  test("bricht nach 50 Seiten mit stabilem Cursor ab und liefert Teilmenge", async () => {
    (global as any).fetch = jest.fn().mockResolvedValue({
      ok: true, status: 200,
      json: async () => ({ results: [{ id: "x", content: "T", project_id: "p" }], next_cursor: "immerGleich" }),
      text: async () => "",
    });
    const errSpy = jest.spyOn(console, "error").mockImplementation(() => {});
    const tasks = await getAllTasks("tok");
    expect((global as any).fetch).toHaveBeenCalledTimes(50);
    expect(tasks).toHaveLength(50);
    expect(errSpy).toHaveBeenCalled();
    errSpy.mockRestore();
  });
});
```

- [ ] **Step 2: Tests laufen lassen, muessen failen**

Run: `npx jest tests/todoist.test.ts`
Expected: FAIL (createTask-Signatur, isAuthError fehlt, maxPages-Test loopt nicht ab bzw. Timeout; Jest-Default-Timeout reicht fuer 50 gemockte Seiten).

- [ ] **Step 3: Implementieren**

In `src/lib/todoist.ts`:

```ts
export interface NewTaskOptions {
  content: string;
  project_id?: string;
  priority?: number;     // 1 bis 4, 4 = hoechste (P1)
  due_date?: string;     // YYYY-MM-DD, deterministisch (Chips)
  due_string?: string;   // natuerliche Sprache, von Todoist geparst
  due_lang?: string;
}

// true nur fuer echte Auth-Probleme (Token ungueltig/gesperrt); alles andere ist retry-wuerdig.
export function isAuthError(e: unknown): boolean {
  return e instanceof TodoistError && (e.status === 401 || e.status === 403);
}
```

`createTask` ersetzen:

```ts
export async function createTask(token: string, options: NewTaskOptions): Promise<TodoistTask> {
  const res = await fetch(`${API}/tasks`, {
    method: "POST",
    headers: { ...auth(token), "Content-Type": "application/json" },
    body: JSON.stringify(options),
  });
  return (await ensureOk(res)).json();
}
```

`getAllTasks` erweitern (Schleifenkopf + Zaehler; Rest unveraendert):

```ts
const MAX_PAGES = 50; // 10'000 Tasks; Endlosschleifen-Versicherung falls die API einen stabilen Cursor liefert

export async function getAllTasks(token: string): Promise<TodoistTask[]> {
  const all: TodoistTask[] = [];
  let cursor: string | null | undefined;
  let pages = 0;
  do {
    const url = `${API}/tasks?limit=200${cursor ? `&cursor=${encodeURIComponent(cursor)}` : ""}`;
    const res = await fetch(url, { headers: auth(token) });
    const data: Paginated<TodoistTask> | TodoistTask[] = await (await ensureOk(res)).json();
    all.push(...unwrap(data));
    cursor = Array.isArray(data) ? null : data.next_cursor;
    pages++;
    if (pages >= MAX_PAGES && cursor) {
      console.error(`getAllTasks: Abbruch nach ${MAX_PAGES} Seiten, Task-Liste evtl. unvollstaendig.`);
      cursor = null;
    }
  } while (cursor);
  return all;
}
```

In `src/taskpane/taskpane.ts` den einzigen createTask-Aufruf (in `wireNewTask`) anpassen:

```ts
const task = await createTask(token, { content: prepared.subject || "Mail" });
```

- [ ] **Step 4: Alle Tests + Build**

Run: `npx jest && npm run build`
Expected: PASS + Build clean (Aufrufer wurde mitgezogen).

- [ ] **Step 5: Commit**

```bash
git add src/lib/todoist.ts src/taskpane/taskpane.ts tests/todoist.test.ts
git commit -m "Todoist: createTask mit NewTaskOptions, maxPages-Guard, isAuthError"
```

---

### Task 3: taskLogic: buildNewTaskOptions

**Files:**
- Modify: `src/taskpane/taskLogic.ts`
- Test: `tests/taskLogic.test.ts`

**Interfaces:**
- Consumes: `NewTaskOptions` aus `../lib/todoist` (Task 2).
- Produces (Task 6 verlaesst sich exakt darauf):
  - `type DueChip = "today" | "tomorrow" | "nextWeek" | "none"`
  - `interface NewTaskInput { title: string; projectId: string | null; priority: number; dueChip: DueChip; dueText: string; today: string }`
  - `buildNewTaskOptions(input: NewTaskInput): NewTaskOptions`

- [ ] **Step 1: Failing Tests schreiben**

In `tests/taskLogic.test.ts` (Import um `buildNewTaskOptions` erweitern):

```ts
describe("buildNewTaskOptions", () => {
  const base = { title: "Task", projectId: null, priority: 1, dueChip: "none" as const, dueText: "", today: "2026-07-06" };

  test("minimal: nur content", () => {
    expect(buildNewTaskOptions(base)).toEqual({ content: "Task" });
  });
  test("Chips mappen deterministisch auf due_date (+0/+1/+7)", () => {
    expect(buildNewTaskOptions({ ...base, dueChip: "today" }).due_date).toBe("2026-07-06");
    expect(buildNewTaskOptions({ ...base, dueChip: "tomorrow" }).due_date).toBe("2026-07-07");
    expect(buildNewTaskOptions({ ...base, dueChip: "nextWeek" }).due_date).toBe("2026-07-13");
  });
  test("Monatswechsel beim Addieren", () => {
    expect(buildNewTaskOptions({ ...base, dueChip: "tomorrow", today: "2026-07-31" }).due_date).toBe("2026-08-01");
  });
  test("Freitext gewinnt ueber Chip und geht als due_string mit de", () => {
    const o = buildNewTaskOptions({ ...base, dueChip: "today", dueText: "naechsten freitag" });
    expect(o.due_string).toBe("naechsten freitag");
    expect(o.due_lang).toBe("de");
    expect(o.due_date).toBeUndefined();
  });
  test("Prioritaet 1 (keine) wird weggelassen, andere gesendet", () => {
    expect(buildNewTaskOptions(base).priority).toBeUndefined();
    expect(buildNewTaskOptions({ ...base, priority: 4 }).priority).toBe(4);
  });
  test("leerer Titel faellt auf Mail zurueck, projectId wird uebernommen", () => {
    const o = buildNewTaskOptions({ ...base, title: "   ", projectId: "p7" });
    expect(o.content).toBe("Mail");
    expect(o.project_id).toBe("p7");
  });
});
```

- [ ] **Step 2: Tests laufen lassen, muessen failen**

Run: `npx jest tests/taskLogic.test.ts`
Expected: FAIL, Export fehlt.

- [ ] **Step 3: Implementieren**

In `src/taskpane/taskLogic.ts` (Import oben um `NewTaskOptions` aus `../lib/todoist` erweitern):

```ts
export type DueChip = "today" | "tomorrow" | "nextWeek" | "none";

export interface NewTaskInput {
  title: string;
  projectId: string | null;
  priority: number; // 1..4 wie Todoist, 1 = keine (P4)
  dueChip: DueChip;
  dueText: string;
  today: string; // YYYY-MM-DD
}

function addDays(iso: string, days: number): string {
  const d = new Date(`${iso}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

// Freitext gewinnt ueber Chip (geht als due_string an Todoists Sprach-Parser);
// Chips erzeugen deterministische due_date ohne Parsing-Risiko.
export function buildNewTaskOptions(input: NewTaskInput): NewTaskOptions {
  const opts: NewTaskOptions = { content: input.title.trim() || "Mail" };
  if (input.projectId) opts.project_id = input.projectId;
  if (input.priority > 1) opts.priority = input.priority;
  const dueText = input.dueText.trim();
  if (dueText) {
    opts.due_string = dueText;
    opts.due_lang = "de";
  } else if (input.dueChip === "today") {
    opts.due_date = input.today;
  } else if (input.dueChip === "tomorrow") {
    opts.due_date = addDays(input.today, 1);
  } else if (input.dueChip === "nextWeek") {
    opts.due_date = addDays(input.today, 7);
  }
  return opts;
}
```

- [ ] **Step 4: Tests laufen lassen, muessen passen**

Run: `npx jest tests/taskLogic.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/taskLogic.ts tests/taskLogic.test.ts
git commit -m "TaskLogic: buildNewTaskOptions (Chips deterministisch, Freitext via due_string)"
```

---

### Task 4: attachToTask: senderName, Kommentartext, prepareMail/readAndPrepareCurrentMail

**Files:**
- Modify: `src/lib/attachToTask.ts`
- Test: `tests/attachToTask.prepare.test.ts` (bestehende commentText-Assertions anpassen!)

**Interfaces:**
- Consumes: `MailData` aus `./emlBuilder` (Felder: subject, from, to, cc, date, htmlBody, attachments), `readCurrentMail` aus `./mailReader`. `MailData.from` ist bereits formatiert als `"Name <adresse>"` oder `"adresse"` oder `""`.
- Produces (Task 7 verlaesst sich exakt darauf):
  - `senderName(from: string): string` (exportiert)
  - `prepareMail(mail: MailData, includeAttachments: boolean): PreparedMail` (exportiert, rein; `false` baut die .eml ohne Datei-Anhaenge)
  - `readAndPrepareCurrentMail(): Promise<{ mail: MailData; prepared: PreparedMail }>` (exportiert)
  - `prepareCurrentMail()` bleibt und delegiert (Rueckgabe unveraendert `PreparedMail`)
  - Kommentartext neu: `Betreff (Datum, von Name)`; fehlt der Name: `Betreff (Datum)`; fehlt das Datum: `Betreff (von Name)`; fehlt beides: `Betreff`.

- [ ] **Step 1: Failing Tests schreiben**

In `tests/attachToTask.prepare.test.ts` (Import um `senderName, prepareMail` erweitern; `MailData` aus `../src/lib/emlBuilder` importieren). Das bestehende `readCurrentMail`-Mock der Datei liefert die Basis-Mail; zusaetzlich eine lokale Fixture bauen:

```ts
const mailFixture = (over: Partial<MailData> = {}): MailData => ({
  subject: "Rechnung Q3",
  from: "Joel Willi <joel@wowoni.ch>",
  to: [],
  cc: [],
  date: "Sun, 05 Jul 2026 10:00:00 GMT",
  htmlBody: "<p>Hallo</p>",
  attachments: [{ name: "a.pdf", contentType: "application/pdf", base64: "QUJD" }],
  ...over,
});

describe("senderName", () => {
  test("Name <adresse> liefert den Namen", () => {
    expect(senderName("Joel Willi <joel@wowoni.ch>")).toBe("Joel Willi");
  });
  test("nur Adresse liefert die Adresse", () => {
    expect(senderName("joel@wowoni.ch")).toBe("joel@wowoni.ch");
  });
  test("leer liefert leer", () => {
    expect(senderName("")).toBe("");
  });
});

describe("prepareMail commentText", () => {
  test("Betreff (Datum, von Name)", () => {
    expect(prepareMail(mailFixture(), true).commentText).toBe("Rechnung Q3 (05.07.2026, von Joel Willi)");
  });
  test("ohne Absender: Betreff (Datum)", () => {
    expect(prepareMail(mailFixture({ from: "" }), true).commentText).toBe("Rechnung Q3 (05.07.2026)");
  });
  test("ohne Datum: Betreff (von Name)", () => {
    expect(prepareMail(mailFixture({ date: "kaputt" }), true).commentText).toBe("Rechnung Q3 (von Joel Willi)");
  });
  test("ohne beides: nur Betreff", () => {
    expect(prepareMail(mailFixture({ from: "", date: "kaputt" }), true).commentText).toBe("Rechnung Q3");
  });
});

describe("prepareMail ohne Anhaenge", () => {
  test("laesst Datei-Anhaenge weg und wird kleiner", async () => {
    const full = prepareMail(mailFixture(), true);
    const textOnly = prepareMail(mailFixture(), false);
    expect(textOnly.sizeBytes).toBeLessThan(full.sizeBytes);
    const eml = await textOnly.blob.text();
    expect(eml).not.toContain("a.pdf");
    expect(eml).toContain("Rechnung Q3");
  });
});
```

WICHTIG: Bestehende Tests in dieser Datei, die `prepareCurrentMail().commentText` als `"Betreff (Datum)"` asserten, muessen auf das neue Format angepasst werden (je nachdem, ob das Mock einen `from` mit Anzeigenamen setzt, kommt `, von ...` dazu). Assertions anfassen, nicht das Verhalten verbiegen.

- [ ] **Step 2: Tests laufen lassen, muessen failen**

Run: `npx jest tests/attachToTask.prepare.test.ts`
Expected: FAIL (Exports fehlen).

- [ ] **Step 3: Implementieren**

In `src/lib/attachToTask.ts`:

```ts
// "Anzeige Name <adresse>" -> "Anzeige Name"; ohne Anzeigename -> Adresse; leer -> "".
export function senderName(from: string): string {
  const m = from.match(/^(.*?)\s*<([^>]*)>$/);
  if (m) return m[1].trim() || m[2].trim();
  return from.trim();
}

function buildCommentText(subject: string, datePart: string, name: string): string {
  const meta = [datePart, name ? `von ${name}` : ""].filter(Boolean).join(", ");
  return meta ? `${subject} (${meta})` : subject;
}

export function prepareMail(mail: MailData, includeAttachments: boolean): PreparedMail {
  const effective = includeAttachments ? mail : { ...mail, attachments: [] };
  const { blob, fileName } = emlBlobFor(effective);
  const commentText = buildCommentText(mail.subject, formatMailDate(mail.date), senderName(mail.from));
  const bodyText = htmlToText(mail.htmlBody).slice(0, 2000);
  return { blob, fileName, sizeBytes: blob.size, subject: mail.subject, commentText, bodyText };
}

export async function readAndPrepareCurrentMail(): Promise<{ mail: MailData; prepared: PreparedMail }> {
  const mail = await readCurrentMail();
  return { mail, prepared: prepareMail(mail, true) };
}

export async function prepareCurrentMail(): Promise<PreparedMail> {
  return (await readAndPrepareCurrentMail()).prepared;
}
```

(Die alte Inline-Logik in `prepareCurrentMail` entfaellt; `attachCurrentMailToTask` bleibt unveraendert.)

- [ ] **Step 4: Alle Tests laufen lassen**

Run: `npx jest`
Expected: PASS (inkl. angepasster Bestands-Assertions; `attachToTask.orchestration.test.ts` laeuft ueber `prepareCurrentMail` und muss ggf. ebenfalls neue commentText-Erwartung bekommen).

- [ ] **Step 5: Commit**

```bash
git add src/lib/attachToTask.ts tests/attachToTask.prepare.test.ts
git commit -m "Ablage-Titel: Betreff (Datum, von Name) + prepareMail mit Text-only-Variante"
```

(Weitere Testdateien nur stagen, wenn tatsaechlich angepasst.)

---

### Task 5: taskpane: Pfeiltasten-Navigation + Retry-Banner

**Files:**
- Modify: `src/taskpane/taskpane.ts`
- Modify: `src/taskpane/taskpane.html` (Fehlerbanner-Markup)
- Modify: `src/taskpane/taskpane.css` (.selected-Stil, #load-error-Stil)

**Interfaces:**
- Consumes: `moveSelection` (Task 1), `isAuthError` (Task 2).
- Produces: `rerender()` in taskpane.ts (Task 7 ruft sie auf): rendert Trefferliste, wenn das Suchfeld Text enthaelt, sonst die Standardansicht.

- [ ] **Step 1: HTML-Banner einfuegen**

In `src/taskpane/taskpane.html` innerhalb `#list-region` direkt nach dem `#skeleton`-Div:

```html
<div id="load-error" hidden>
  <p id="load-error-text"></p>
  <button id="load-error-retry" class="primary">Erneut versuchen</button>
</div>
```

- [ ] **Step 2: CSS ergaenzen**

In `src/taskpane/taskpane.css` (an die bestehenden Variablen anlehnen; `--accent` existiert):

```css
.task-row.selected { background: rgba(228, 67, 50, 0.10); box-shadow: inset 2px 0 0 var(--accent); }
@media (prefers-color-scheme: dark) {
  .task-row.selected { background: rgba(228, 67, 50, 0.18); }
}
#load-error { display: flex; flex-direction: column; gap: 8px; align-items: flex-start; padding: 8px 0; }
#load-error p { margin: 0; }
```

- [ ] **Step 3: Auswahl-Logik in taskpane.ts**

Import um `moveSelection` (aus `./taskLogic`) und `isAuthError` (aus `../lib/todoist`) erweitern. Neuer Modul-State + Helfer:

```ts
let selectedIndex = 0;

function visibleRows(): HTMLButtonElement[] {
  return Array.from($("task-groups").querySelectorAll<HTMLButtonElement>(".task-row"));
}

function applySelection(): void {
  const rows = visibleRows();
  rows.forEach((r, i) => r.classList.toggle("selected", i === selectedIndex));
  const sel = rows[selectedIndex];
  if (sel) sel.scrollIntoView({ block: "nearest" });
}
```

Am Ende von `renderSections` (im Nicht-leer-Zweig, nach der for-Schleife):

```ts
selectedIndex = 0;
applySelection();
```

Den bestehenden keydown-Handler in `wireSearch` ersetzen:

```ts
($("search") as HTMLInputElement).addEventListener("keydown", (e) => {
  const key = (e as KeyboardEvent).key;
  if (key === "ArrowDown" || key === "ArrowUp") {
    e.preventDefault();
    selectedIndex = moveSelection(selectedIndex, key === "ArrowDown" ? 1 : -1, visibleRows().length);
    applySelection();
    return;
  }
  if (key !== "Enter") return;
  const rows = visibleRows();
  const target = rows[selectedIndex] ?? rows[0];
  if (target && !target.disabled) target.click();
});
```

- [ ] **Step 4: rerender() + Retry-Weiche**

Neue Funktion (ersetzt direkte renderDefaultView/renderSections-Aufrufe der Such- und Ladepfade):

```ts
// Rendert passend zum aktuellen Suchfeld-Inhalt. Fixt nebenbei: Tippen waehrend
// des Initial-Loads wird beim Load-Ende nicht mehr von der Standardansicht ueberschrieben.
function rerender(): void {
  const q = ($("search") as HTMLInputElement).value.trim();
  if (q) renderSections([["Treffer", filterTasks(allTasks, q, projectNames)]], "Keine Treffer.");
  else renderDefaultView();
}
```

Im Such-Input-Handler den Rumpf ersetzen durch `rerender();` (die q-Logik lebt jetzt dort).
In `loadTasks` den Erfolgs-Zweig `renderDefaultView()` durch `rerender()` ersetzen und den catch-Zweig ersetzen:

```ts
} catch (e) {
  setSkeleton(false);
  if (isAuthError(e)) {
    setStatus(`Token ungültig: ${(e as Error).message}`, "err", e);
    showTokenSection();
  } else {
    console.error("Task-Load fehlgeschlagen", e);
    $("load-error-text").textContent = `Abruf fehlgeschlagen: ${(e as Error).message}`;
    $("load-error").hidden = false;
  }
}
```

Am Anfang von `loadTasks` (vor setSkeleton): `$("load-error").hidden = true;`

Retry-Button einmalig verdrahten (in `wireSearch` oder eigenem `wireRetry()`, aufgerufen neben wireSearch in `start`):

```ts
function wireRetry(): void {
  ($("load-error-retry") as HTMLButtonElement).onclick = () => {
    const token = getToken();
    if (token) void loadTasks(token);
  };
}
```

- [ ] **Step 5: Verifizieren + Commit**

Run: `npx jest && npm run build`
Expected: PASS + Build clean.

```bash
git add src/taskpane/taskpane.ts src/taskpane/taskpane.html src/taskpane/taskpane.css
git commit -m "Task-Pane: Pfeiltasten-Navigation mit Markierung + Retry-Banner statt Token-Screen"
```

---

### Task 6: taskpane: Neues-Task-Formular

**Files:**
- Modify: `src/taskpane/taskpane.html` (Formular nach dem new-task-Button)
- Modify: `src/taskpane/taskpane.css` (Formular-Stile)
- Modify: `src/taskpane/taskpane.ts` (`wireNewTask` ersetzen)

**Interfaces:**
- Consumes: `buildNewTaskOptions`, `DueChip`, `todayIso`, `priorityColor` (taskLogic), `createTask` mit `NewTaskOptions` (Task 2), `attachPreparedToTask`, State `prepared`, `projectNames`, `busy`, `tooLarge()`, `loadTasks`.
- Produces: fertiges Formular-Verhalten; keine neuen Exports.

- [ ] **Step 1: HTML**

In `src/taskpane/taskpane.html` direkt nach dem `#new-task`-Button:

```html
<form id="new-task-form" hidden>
  <input id="nt-title" type="text" placeholder="Task-Titel" autocomplete="off" />
  <select id="nt-project" aria-label="Projekt"></select>
  <div class="nt-row">
    <span class="nt-label">Priorität</span>
    <span class="nt-prio" id="nt-prio">
      <button type="button" data-p="4" title="P1"></button>
      <button type="button" data-p="3" title="P2"></button>
      <button type="button" data-p="2" title="P3"></button>
      <button type="button" data-p="1" title="Keine" class="active"></button>
    </span>
  </div>
  <div class="nt-chips" id="nt-chips">
    <button type="button" data-chip="none" class="active">Kein Datum</button>
    <button type="button" data-chip="today">Heute</button>
    <button type="button" data-chip="tomorrow">Morgen</button>
    <button type="button" data-chip="nextWeek">Nächste Woche</button>
  </div>
  <input id="nt-due-text" type="text" placeholder="oder Fälligkeit als Text, z.B. freitag" autocomplete="off" />
  <div class="nt-actions">
    <button type="submit" id="nt-create" class="primary">Erstellen + Mail anhängen</button>
    <button type="button" id="nt-cancel" class="ghost">Abbrechen</button>
  </div>
</form>
```

- [ ] **Step 2: CSS**

In `src/taskpane/taskpane.css` (Input-/Select-Grundstil an die bestehende `#token-input`-Regel anlehnen, gleiche Border/Radius/Farben-Variablen verwenden):

```css
#new-task-form { display: flex; flex-direction: column; gap: 8px; margin-top: 8px; }
.nt-row { display: flex; align-items: center; gap: 8px; }
.nt-label { font-size: 12px; color: var(--muted, #808080); }
.nt-prio { display: flex; gap: 6px; }
.nt-prio button { width: 20px; height: 20px; border-radius: 50%; border: 2px solid transparent; cursor: pointer; padding: 0; }
.nt-prio button.active { border-color: var(--text, #202020); }
.nt-chips { display: flex; gap: 6px; flex-wrap: wrap; }
.nt-chips button { padding: 4px 10px; border-radius: 14px; border: 1px solid var(--border, #ccc); background: transparent; color: inherit; cursor: pointer; font-size: 12px; }
.nt-chips button.active { background: var(--accent); color: #fff; border-color: var(--accent); }
.nt-actions { display: flex; gap: 8px; }
```

(Existieren `--muted`/`--border`/`--text` nicht unter diesen Namen, die tatsaechlichen Variablennamen aus dem CSS-Kopf der Datei verwenden; Fallback-Werte stehen dahinter.)

- [ ] **Step 3: wireNewTask ersetzen**

In `src/taskpane/taskpane.ts` (Imports um `buildNewTaskOptions, DueChip` aus `./taskLogic` erweitern):

```ts
let ntPriority = 1;
let ntChip: DueChip = "none";

function markActive(container: HTMLElement, active: HTMLElement): void {
  for (const el of Array.from(container.children)) el.classList.toggle("active", el === active);
}

function fillProjectSelect(): void {
  const sel = $("nt-project") as HTMLSelectElement;
  sel.innerHTML = "";
  const entries = Object.entries(projectNames);
  const inboxIdx = entries.findIndex(([, name]) => name === "Inbox" || name === "Eingang");
  if (inboxIdx > 0) entries.unshift(entries.splice(inboxIdx, 1)[0]);
  for (const [id, name] of entries) {
    const o = document.createElement("option");
    o.value = id;
    o.textContent = name;
    sel.appendChild(o);
  }
}

function wireNewTask(): void {
  const btn = $("new-task") as HTMLButtonElement;
  const form = $("new-task-form") as HTMLFormElement;
  btn.hidden = false;

  btn.onclick = () => {
    if (tooLarge() || !prepared) return;
    form.hidden = !form.hidden;
    if (!form.hidden) {
      ($("nt-title") as HTMLInputElement).value = prepared.subject || "";
      fillProjectSelect();
      ($("nt-title") as HTMLInputElement).focus();
    }
  };

  ($("nt-cancel") as HTMLButtonElement).onclick = () => { form.hidden = true; };

  for (const b of Array.from($("nt-prio").querySelectorAll<HTMLButtonElement>("button"))) {
    b.style.background = priorityColor(Number(b.dataset.p));
    b.onclick = () => { ntPriority = Number(b.dataset.p); markActive($("nt-prio"), b); };
  }
  for (const b of Array.from($("nt-chips").querySelectorAll<HTMLButtonElement>("button"))) {
    b.onclick = () => { ntChip = b.dataset.chip as DueChip; markActive($("nt-chips"), b); };
  }

  form.onsubmit = async (e) => {
    e.preventDefault();
    if (busy || tooLarge() || !prepared) return;
    const token = getToken();
    if (!token) { showTokenSection(); return; }
    busy = true;
    ($("nt-create") as HTMLButtonElement).disabled = true;
    setStatus("Erstelle Task...", "");
    try {
      const opts = buildNewTaskOptions({
        title: ($("nt-title") as HTMLInputElement).value,
        projectId: ($("nt-project") as HTMLSelectElement).value || null,
        priority: ntPriority,
        dueChip: ntChip,
        dueText: ($("nt-due-text") as HTMLInputElement).value,
        today: todayIso(new Date()),
      });
      const task = await createTask(token, opts);
      await attachPreparedToTask(token, task.id, prepared);
      setStatus(`Neuer Task "${task.content}" mit Mail erstellt.`, "ok");
      form.hidden = true;
      await loadTasks(token);
    } catch (e2) {
      setStatus(`Fehler beim Erstellen: ${(e2 as Error).message}`, "err", e2); // Formular bleibt offen, Eingaben bleiben
    } finally {
      busy = false;
      ($("nt-create") as HTMLButtonElement).disabled = false;
    }
  };
}
```

- [ ] **Step 4: Verifizieren + Commit**

Run: `npx jest && npm run build`
Expected: PASS + Build clean.

```bash
git add src/taskpane/taskpane.ts src/taskpane/taskpane.html src/taskpane/taskpane.css
git commit -m "Task-Pane: Neues-Task-Formular mit Projekt, Prioritaet und Faelligkeits-Chips"
```

---

### Task 7: taskpane: Text-only-Flow bei >25 MB + Doku

**Files:**
- Modify: `src/taskpane/taskpane.ts`
- Modify: `CLAUDE.md`

**Interfaces:**
- Consumes: `prepareMail`, `readAndPrepareCurrentMail`, `MAX_BYTES` (aus `../lib/attachToTask`, Task 4), `MailData` (aus `../lib/emlBuilder`), `rerender()` (Task 5).
- Produces: fertiges Verhalten; keine neuen Exports.

- [ ] **Step 1: Mail-State + Warnung mit Aktions-Button**

In `src/taskpane/taskpane.ts`: Import anpassen auf
`import { readAndPrepareCurrentMail, prepareMail, attachPreparedToTask, PreparedMail, MAX_BYTES } from "../lib/attachToTask";`
und `import { MailData } from "../lib/emlBuilder";`. Neuer Modul-State:

```ts
let mailData: MailData | null = null;
```

In `start()` den prepare-Block ersetzen:

```ts
try {
  const { mail, prepared: p } = await readAndPrepareCurrentMail();
  mailData = mail;
  prepared = p;
  renderContextBar();
  renderSizeWarning();
} catch (e) {
  setStatus(`Mail konnte nicht gelesen werden: ${(e as Error).message}`, "err", e);
}
```

`renderSizeWarning` ersetzen (nutzt jetzt MAX_BYTES statt lokaler Konstante):

```ts
function renderSizeWarning(): void {
  const warn = $("size-warning");
  warn.innerHTML = "";
  if (prepared && prepared.sizeBytes > MAX_BYTES) {
    const text = document.createElement("span");
    text.textContent = `Mail ist ${(prepared.sizeBytes / 1024 / 1024).toFixed(1)} MB gross (Todoist erlaubt max 25 MB). `;
    warn.appendChild(text);
    if (mailData && mailData.attachments.length > 0) {
      const strip = document.createElement("button");
      strip.className = "ghost";
      strip.textContent = "Ohne Anhänge anhängen";
      strip.onclick = () => {
        prepared = prepareMail(mailData!, false);
        if (prepared.sizeBytes > MAX_BYTES) {
          // Body allein schon zu gross: Warnung aktualisieren, KEIN renderSizeWarning()
          // (das wuerde den wirkungslosen Button erneut anbieten).
          warn.textContent = `Auch ohne Anhänge ist die Mail ${(prepared.sizeBytes / 1024 / 1024).toFixed(1)} MB gross, Anhängen bleibt deaktiviert.`;
          warn.hidden = false;
        } else {
          // Erfolgszweig: Hinweis setzen, NICHT renderSizeWarning() aufrufen
          // (tooLarge() ist jetzt false und wuerde den Hinweis verstecken).
          warn.textContent = "Anhänge werden weggelassen.";
          warn.hidden = false;
          ($("new-task") as HTMLButtonElement).disabled = false;
        }
        rerender(); // Zeilen mit neuem tooLarge()-Zustand neu aufbauen
      };
      warn.appendChild(strip);
    }
    warn.hidden = false;
  } else {
    warn.hidden = true;
  }
  ($("new-task") as HTMLButtonElement).disabled = tooLarge();
}
```

`tooLarge()` auf MAX_BYTES umstellen:

```ts
function tooLarge(): boolean {
  return !!prepared && prepared.sizeBytes > MAX_BYTES;
}
```

- [ ] **Step 2: Verifizieren**

Run: `npx jest && npm run build`
Expected: PASS + Build clean. Zusaetzlich `grep -n "25 \* 1024" src/taskpane/taskpane.ts` -> keine Treffer mehr (nur noch MAX_BYTES).

- [ ] **Step 3: CLAUDE.md nachfuehren**

Modulzeilen aktualisieren (todoist.ts: createTask mit NewTaskOptions, isAuthError, maxPages; taskLogic: moveSelection/buildNewTaskOptions/#-Syntax; attachToTask: senderName/prepareMail/readAndPrepareCurrentMail; taskpane: Pfeiltasten, Formular, Retry-Banner, Text-only-Button) und neuen Status-Eintrag unter "Status / Gotchas" mit Datum 2026-07-06 ergaenzen (Verbesserungspaket: die 5 Features in je einem Halbsatz, neue Testzahl aus dem echten `npx jest`-Lauf).

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.ts CLAUDE.md
git commit -m "Task-Pane: Text-only-Anhaengen bei uebergrossen Mails + Doku-Stand"
```

---

## Verifikation nach Abschluss

- `npx jest` alle gruen, `npm run build` clean.
- Manuell (Manuel, on-device): (1) ↑/↓ + Enter im Suchfeld, (2) `#projekt wort`-Suche, (3) Neuen Task mit Projekt/P2/Morgen erstellen, (4) Ablage-Kommentar zeigt "von Name", (5) Riesen-Mail: Button "Ohne Anhänge anhängen".
- Deploy (`npm run deploy`) erst nach Manuels Okay; vorher `gh auth status` (Account manuelweingartner), danach zurueck auf CMI-Kunden.
