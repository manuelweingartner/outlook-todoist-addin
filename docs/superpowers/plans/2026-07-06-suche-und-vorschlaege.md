# Bessere Suche + Top-3-Vorschläge Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Case-insensitive Client-Suche über alle offenen Todoist-Tasks plus eine "Vorschläge"-Sektion mit den Top-3-Tasks, die zum Betreff/Inhalt der offenen Mail passen.

**Architecture:** Beim Pane-Start werden einmal ALLE offenen Tasks paginiert geladen (`getAllTasks`, folgt `next_cursor`). Suche und Vorschläge laufen danach rein client-seitig über neue reine Funktionen in `taskLogic.ts`. Die server-seitige Suche (`searchTasks`) und der Server-Filter (`getTasks`) entfallen.

**Tech Stack:** TypeScript, Office.js, Todoist Unified API v1, Jest (ts-jest, jsdom), Webpack.

**Spec:** `docs/superpowers/specs/2026-07-06-suche-und-vorschlaege-design.md`

## Global Constraints

- Git-Identität: Manuel Weingartner <manuel.weingartner@gmx.ch>. NIE Co-Authored-By.
- Keine Em-/En-Dashes in Code, Kommentaren, Docs.
- Echte Umlaute (ä/ö/ü) in sichtbarem Deutsch; Schweizer Deutsch, kein ß (gross/ausser).
- Kein stilles try/catch: Fehler loggen UND dem Nutzer zeigen.
- TDD: Test zuerst, dann Implementation.
- Tests laufen mit `npx jest`, Build mit `npm run build`.
- Kein `git add -A`: Dateien immer explizit stagen.

---

### Task 1: `getAllTasks` mit Cursor-Pagination (todoist.ts)

**Files:**
- Modify: `src/lib/todoist.ts` (nach `getTasks`, Zeile ~53)
- Test: `tests/todoist.test.ts`

**Interfaces:**
- Consumes: bestehende Helfer `auth(token)`, `ensureOk(res)`, `unwrap(data)`, `Paginated<T>`, `API`-Konstante in `src/lib/todoist.ts`.
- Produces: `getAllTasks(token: string): Promise<TodoistTask[]>` (lädt ALLE offenen Tasks über alle Seiten). Task 5 benutzt genau diese Signatur.

- [ ] **Step 1: Failing Test schreiben**

In `tests/todoist.test.ts` (Import oben um `getAllTasks` erweitern):

```ts
describe("getAllTasks", () => {
  function pageMock(pages: Array<{ results: unknown[]; next_cursor: string | null }>) {
    const fn = jest.fn();
    for (const p of pages) {
      fn.mockResolvedValueOnce({ ok: true, status: 200, json: async () => p, text: async () => "" });
    }
    (global as any).fetch = fn;
    return fn;
  }

  test("laedt eine Einzelseite", async () => {
    const fn = pageMock([{ results: [{ id: "1", content: "A", project_id: "p" }], next_cursor: null }]);
    const tasks = await getAllTasks("tok");
    expect(tasks.map((t) => t.id)).toEqual(["1"]);
    expect(fn).toHaveBeenCalledTimes(1);
    const [url, opts] = fn.mock.calls[0];
    expect(url).toBe("https://api.todoist.com/api/v1/tasks?limit=200");
    expect(opts.headers.Authorization).toBe("Bearer tok");
  });

  test("folgt next_cursor ueber mehrere Seiten und konkateniert", async () => {
    const fn = pageMock([
      { results: [{ id: "1", content: "A", project_id: "p" }], next_cursor: "c2" },
      { results: [{ id: "2", content: "B", project_id: "p" }], next_cursor: null },
    ]);
    const tasks = await getAllTasks("tok");
    expect(tasks.map((t) => t.id)).toEqual(["1", "2"]);
    expect(fn).toHaveBeenCalledTimes(2);
    expect(fn.mock.calls[1][0]).toBe("https://api.todoist.com/api/v1/tasks?limit=200&cursor=c2");
  });

  test("wirft TodoistError bei 401", async () => {
    (global as any).fetch = jest.fn().mockResolvedValue({
      ok: false, status: 401, json: async () => ({}), text: async () => "unauthorized",
    });
    await expect(getAllTasks("bad")).rejects.toBeInstanceOf(TodoistError);
  });
});
```

- [ ] **Step 2: Test laufen lassen, muss failen**

Run: `npx jest tests/todoist.test.ts`
Expected: FAIL, `getAllTasks` ist kein Export.

- [ ] **Step 3: Minimal implementieren**

In `src/lib/todoist.ts` nach `getTasks` einfügen:

```ts
// Laedt ALLE offenen Tasks: v1 paginiert mit max. 200 pro Seite, wir folgen
// next_cursor bis zum Ende. Basis fuer client-seitige Suche + Vorschlaege.
export async function getAllTasks(token: string): Promise<TodoistTask[]> {
  const all: TodoistTask[] = [];
  let cursor: string | null | undefined;
  do {
    const url = `${API}/tasks?limit=200${cursor ? `&cursor=${encodeURIComponent(cursor)}` : ""}`;
    const res = await fetch(url, { headers: auth(token) });
    const data: Paginated<TodoistTask> | TodoistTask[] = await (await ensureOk(res)).json();
    all.push(...unwrap(data));
    cursor = Array.isArray(data) ? null : data.next_cursor;
  } while (cursor);
  return all;
}
```

- [ ] **Step 4: Tests laufen lassen, muss passen**

Run: `npx jest tests/todoist.test.ts`
Expected: PASS (alle, inkl. bestehender).

- [ ] **Step 5: Commit**

```bash
git add src/lib/todoist.ts tests/todoist.test.ts
git commit -m "Todoist: getAllTasks laedt alle offenen Tasks ueber next_cursor-Pagination"
```

---

### Task 2: `filterTasks` + `dueTodayOrOverdue` (taskLogic.ts)

**Files:**
- Modify: `src/taskpane/taskLogic.ts`
- Test: `tests/taskLogic.test.ts`

**Interfaces:**
- Consumes: `TodoistTask` aus `../lib/todoist`.
- Produces:
  - `filterTasks(tasks: TodoistTask[], query: string, projectNames: Record<string, string>): TodoistTask[]` (case-insensitiv, Mehrwort-UND, matcht Titel + Projektname; leere Query = alle Tasks).
  - `dueTodayOrOverdue(tasks: TodoistTask[], today: string): TodoistTask[]` (nur Tasks mit Fälligkeit <= heute; repliziert den bisherigen Server-Filter `(today | overdue)`).

- [ ] **Step 1: Failing Tests schreiben**

In `tests/taskLogic.test.ts` (Import oben um `filterTasks, dueTodayOrOverdue` erweitern; der Helper `t(...)` existiert dort schon):

```ts
const named = (id: string, content: string, project_id = "p1"): TodoistTask => ({
  id, content, project_id, priority: 1, due: null,
});

describe("filterTasks", () => {
  const projects = { p1: "Inbox", p2: "SAP" };

  test("matcht case-insensitiv: 'sap' findet 'SAP'", () => {
    const tasks = [named("a", "SAP Lizenzen klären"), named("b", "Zmittag buchen")];
    expect(filterTasks(tasks, "sap", projects).map((x) => x.id)).toEqual(["a"]);
  });

  test("mehrere Woerter sind UND-verknuepft", () => {
    const tasks = [named("a", "SAP Lizenzen klären"), named("b", "SAP Schulung planen")];
    expect(filterTasks(tasks, "sap schulung", projects).map((x) => x.id)).toEqual(["b"]);
  });

  test("matcht auch den Projektnamen", () => {
    const tasks = [named("a", "Rechnung prüfen", "p2"), named("b", "Rechnung zahlen", "p1")];
    expect(filterTasks(tasks, "sap", projects).map((x) => x.id)).toEqual(["a"]);
  });

  test("leere Query liefert alle Tasks", () => {
    const tasks = [named("a", "x"), named("b", "y")];
    expect(filterTasks(tasks, "  ", projects)).toHaveLength(2);
  });

  test("unbekannte project_id crasht nicht", () => {
    const tasks = [named("a", "SAP Thema", "p99")];
    expect(filterTasks(tasks, "sap", projects).map((x) => x.id)).toEqual(["a"]);
  });
});

describe("dueTodayOrOverdue", () => {
  test("behaelt heute + ueberfaellig, wirft ohne Datum und Zukunft raus", () => {
    const tasks = [t("alt", "2026-07-01"), t("heute", "2026-07-06"), t("morgen", "2026-07-07"), t("ohne", null)];
    expect(dueTodayOrOverdue(tasks, "2026-07-06").map((x) => x.id)).toEqual(["alt", "heute"]);
  });
});
```

- [ ] **Step 2: Test laufen lassen, muss failen**

Run: `npx jest tests/taskLogic.test.ts`
Expected: FAIL, Exports fehlen.

- [ ] **Step 3: Minimal implementieren**

In `src/taskpane/taskLogic.ts` anhängen:

```ts
// Client-seitige Suche: case-insensitiv, alle Woerter muessen matchen (UND),
// gesucht wird in Task-Titel UND Projektname ("sap" findet Projekt "SAP").
export function filterTasks(
  tasks: TodoistTask[],
  query: string,
  projectNames: Record<string, string>,
): TodoistTask[] {
  const words = query.toLowerCase().split(/\s+/).filter(Boolean);
  if (words.length === 0) return tasks;
  return tasks.filter((task) => {
    const hay = `${task.content} ${projectNames[task.project_id] ?? ""}`.toLowerCase();
    return words.every((w) => hay.includes(w));
  });
}

// Repliziert den frueheren Server-Filter "(today | overdue)" client-seitig.
export function dueTodayOrOverdue(tasks: TodoistTask[], today: string): TodoistTask[] {
  return tasks.filter((task) => !!task.due?.date && task.due.date <= today);
}
```

- [ ] **Step 4: Tests laufen lassen, muss passen**

Run: `npx jest tests/taskLogic.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/taskLogic.ts tests/taskLogic.test.ts
git commit -m "TaskLogic: filterTasks (case-insensitiv, Titel+Projekt) + dueTodayOrOverdue"
```

---

### Task 3: `extractMailKeywords` + `suggestTasks` (taskLogic.ts)

**Files:**
- Modify: `src/taskpane/taskLogic.ts`
- Test: `tests/taskLogic.test.ts`

**Interfaces:**
- Consumes: `TodoistTask` aus `../lib/todoist`.
- Produces:
  - `interface MailKeywords { subjectWords: string[]; bodyWords: string[] }`
  - `extractMailKeywords(subject: string, bodyText: string): MailKeywords`
  - `suggestTasks(tasks: TodoistTask[], kw: MailKeywords): TodoistTask[]` (max. 3, Score > 0, Betreff-Wort = 3 Punkte, Body-Wort = 1 Punkt).

- [ ] **Step 1: Failing Tests schreiben**

In `tests/taskLogic.test.ts` (Import um `extractMailKeywords, suggestTasks` erweitern; Helper `named` aus Task 2 wiederverwenden):

```ts
describe("extractMailKeywords", () => {
  test("entfernt gestapelte RE:/AW:/FW:/WG:-Praefixe", () => {
    const kw = extractMailKeywords("AW: RE: SAP Lizenzverlängerung", "");
    expect(kw.subjectWords).toEqual(["sap", "lizenzverlängerung"]);
  });

  test("filtert Stoppwoerter und Kurzwoerter", () => {
    const kw = extractMailKeywords("Die Offerte für das neue CRM", "");
    expect(kw.subjectWords).toEqual(["offerte", "neue", "crm"]);
  });

  test("Body-Woerter dedupliziert und ohne Betreff-Duplikate", () => {
    const kw = extractMailKeywords("SAP Update", "Das SAP Update betrifft die Migration. Migration startet bald.");
    expect(kw.subjectWords).toEqual(["sap", "update"]);
    expect(kw.bodyWords).toEqual(["betrifft", "migration", "startet", "bald"]);
  });

  test("Body wird bei 2000 Zeichen gekappt", () => {
    const kw = extractMailKeywords("x", "a".repeat(2000) + " spätwort");
    expect(kw.bodyWords).not.toContain("spätwort");
  });

  test("leerer Input liefert leere Listen", () => {
    expect(extractMailKeywords("", "")).toEqual({ subjectWords: [], bodyWords: [] });
  });
});

describe("suggestTasks", () => {
  const kw = { subjectWords: ["sap", "lizenz"], bodyWords: ["migration"] };

  test("gewichtet Betreff 3, Body 1 und sortiert absteigend", () => {
    const tasks = [
      named("nurBody", "Migration vorbereiten"),
      named("beide", "SAP Lizenz verlängern"),
      named("einSubject", "SAP Schulung"),
    ];
    expect(suggestTasks(tasks, kw).map((x) => x.id)).toEqual(["beide", "einSubject", "nurBody"]);
  });

  test("Score 0 wird nie vorgeschlagen", () => {
    expect(suggestTasks([named("a", "Zmittag buchen")], kw)).toEqual([]);
  });

  test("maximal 3 Vorschlaege", () => {
    const tasks = ["a", "b", "c", "d"].map((id) => named(id, `SAP ${id}`));
    expect(suggestTasks(tasks, kw)).toHaveLength(3);
  });

  test("matcht case-insensitiv gegen den Task-Titel", () => {
    expect(suggestTasks([named("a", "sap lizenzen")], kw).map((x) => x.id)).toEqual(["a"]);
  });

  test("stabile Reihenfolge bei Gleichstand", () => {
    const tasks = [named("erst", "SAP eins"), named("zweit", "SAP zwei")];
    expect(suggestTasks(tasks, kw).map((x) => x.id)).toEqual(["erst", "zweit"]);
  });
});
```

- [ ] **Step 2: Test laufen lassen, muss failen**

Run: `npx jest tests/taskLogic.test.ts`
Expected: FAIL, Exports fehlen.

- [ ] **Step 3: Minimal implementieren**

In `src/taskpane/taskLogic.ts` anhängen:

```ts
// Kleine DE/EN-Stoppwortliste fuer das Vorschlags-Scoring. Bewusst kurz:
// falsche Positives kosten nur einen schwachen Vorschlag, keine Fehler.
const STOPWORDS = new Set([
  "der", "die", "das", "den", "dem", "des", "ein", "eine", "einen", "einem", "einer",
  "und", "oder", "aber", "nicht", "mit", "von", "für", "auf", "aus", "bei", "nach",
  "über", "unter", "vor", "zum", "zur", "ist", "sind", "war", "wird", "werden",
  "wurde", "hat", "haben", "kann", "können", "muss", "müssen", "soll", "sich",
  "auch", "noch", "nur", "schon", "wie", "wir", "ihr", "sie", "ich",
  "the", "and", "for", "are", "but", "not", "with", "from", "this", "that", "you",
  "your", "have", "has", "was", "were", "will", "would", "can", "could", "should",
  "our", "all", "any", "been", "they", "them", "their",
  "hallo", "liebe", "lieber", "grüsse", "gruss", "freundliche", "freundlichen",
  "danke", "mail", "regards", "dear", "hello", "thanks",
]);

export interface MailKeywords { subjectWords: string[]; bodyWords: string[]; }

function tokenize(text: string): string[] {
  return text
    .toLowerCase()
    .split(/[^\p{L}\p{N}]+/u)
    .filter((w) => w.length >= 3 && !STOPWORDS.has(w));
}

export function extractMailKeywords(subject: string, bodyText: string): MailKeywords {
  const cleanSubject = subject.replace(/^((re|aw|fw|fwd|wg)\s*:\s*)+/i, "");
  const subjectWords = [...new Set(tokenize(cleanSubject))];
  const subjectSet = new Set(subjectWords);
  const bodyWords = [...new Set(tokenize(bodyText.slice(0, 2000)))].filter((w) => !subjectSet.has(w));
  return { subjectWords, bodyWords };
}

// Top-3-Vorschlaege: Betreff-Wort im Titel = 3 Punkte, Body-Wort = 1 Punkt.
// Array.prototype.sort ist stabil, Gleichstand behaelt also die Task-Reihenfolge.
export function suggestTasks(tasks: TodoistTask[], kw: MailKeywords): TodoistTask[] {
  const scored = tasks
    .map((task) => {
      const hay = task.content.toLowerCase();
      let score = 0;
      for (const w of kw.subjectWords) if (hay.includes(w)) score += 3;
      for (const w of kw.bodyWords) if (hay.includes(w)) score += 1;
      return { task, score };
    })
    .filter((s) => s.score > 0);
  scored.sort((a, b) => b.score - a.score);
  return scored.slice(0, 3).map((s) => s.task);
}
```

- [ ] **Step 4: Tests laufen lassen, muss passen**

Run: `npx jest tests/taskLogic.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/taskLogic.ts tests/taskLogic.test.ts
git commit -m "TaskLogic: Mail-Keywords (Praefix/Stoppwort-Filter) + Top-3-Task-Scoring"
```

---

### Task 4: `htmlToText` + `PreparedMail.bodyText` (attachToTask.ts)

**Files:**
- Modify: `src/lib/attachToTask.ts`
- Test: `tests/attachToTask.prepare.test.ts`

**Interfaces:**
- Consumes: `MailData.htmlBody` (bereits von `readCurrentMail` geliefert).
- Produces:
  - `htmlToText(html: string): string` (exportiert, rein, regex-basiert).
  - `PreparedMail` erhält neues Pflichtfeld `bodyText: string` (Klartext, max. 2000 Zeichen). Task 5 liest `prepared.bodyText`.

- [ ] **Step 1: Failing Tests schreiben**

In `tests/attachToTask.prepare.test.ts` (Import um `htmlToText` erweitern; die Datei mockt `readCurrentMail` bereits für `prepareCurrentMail`-Tests, bestehendes Mock-Muster übernehmen):

```ts
describe("htmlToText", () => {
  test("strippt Tags und kollabiert Whitespace", () => {
    expect(htmlToText("<div><p>Hallo  <b>Welt</b></p>\n<p>Zeile 2</p></div>")).toBe("Hallo Welt Zeile 2");
  });

  test("entfernt style- und script-Bloecke samt Inhalt", () => {
    expect(htmlToText("<style>p{color:red}</style><script>var x=1;</script>Text")).toBe("Text");
  });

  test("dekodiert gaengige Entities", () => {
    expect(htmlToText("A&nbsp;&amp;&nbsp;B &lt;C&gt; &quot;D&quot;")).toBe('A & B <C> "D"');
  });
});

describe("prepareCurrentMail bodyText", () => {
  test("liefert gestrippten Body, gekappt auf 2000 Zeichen", async () => {
    // Bestehendes readCurrentMail-Mock der Datei nutzen und htmlBody setzen auf:
    // "<p>" + "a".repeat(3000) + "</p>"
    const prepared = await prepareCurrentMail();
    expect(prepared.bodyText).toHaveLength(2000);
    expect(prepared.bodyText.startsWith("aaa")).toBe(true);
  });
});
```

- [ ] **Step 2: Test laufen lassen, muss failen**

Run: `npx jest tests/attachToTask.prepare.test.ts`
Expected: FAIL, `htmlToText` fehlt / `bodyText` undefined.

- [ ] **Step 3: Minimal implementieren**

In `src/lib/attachToTask.ts`:

```ts
// PreparedMail erweitern:
export interface PreparedMail {
  blob: Blob;
  fileName: string;
  sizeBytes: number;
  subject: string;
  commentText: string;
  bodyText: string;
}

// Regex-Strip reicht fuer Scoring-Zwecke (kein DOM noetig, laeuft auch in Tests ohne jsdom-Kosten).
export function htmlToText(html: string): string {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#\d+;/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}
```

In `prepareCurrentMail` das Rückgabeobjekt erweitern:

```ts
export async function prepareCurrentMail(): Promise<PreparedMail> {
  const mail = await readCurrentMail();
  const { blob, fileName } = emlBlobFor(mail);
  const datePart = formatMailDate(mail.date);
  const commentText = datePart ? `${mail.subject} (${datePart})` : mail.subject;
  const bodyText = htmlToText(mail.htmlBody).slice(0, 2000);
  return { blob, fileName, sizeBytes: blob.size, subject: mail.subject, commentText, bodyText };
}
```

- [ ] **Step 4: Alle Tests laufen lassen**

Run: `npx jest`
Expected: PASS. Achtung: andere Tests, die `PreparedMail`-Objekte literal bauen (z.B. `attachToTask.orchestration.test.ts`), brauchen jetzt ein `bodyText: ""`-Feld; falls TypeScript dort meckert, ergänzen.

- [ ] **Step 5: Commit**

```bash
git add src/lib/attachToTask.ts tests/attachToTask.prepare.test.ts tests/attachToTask.orchestration.test.ts
git commit -m "Mail-Body als Klartext: htmlToText + PreparedMail.bodyText (max 2000 Zeichen)"
```

(Zweite Testdatei nur stagen, wenn sie tatsächlich angepasst wurde.)

---

### Task 5: UI-Verdrahtung: Voll-Load, Client-Suche, Vorschläge-Sektion (taskpane.ts)

**Files:**
- Modify: `src/taskpane/taskpane.ts`

**Interfaces:**
- Consumes: `getAllTasks` (Task 1), `filterTasks`, `dueTodayOrOverdue` (Task 2), `extractMailKeywords`, `suggestTasks` (Task 3), `prepared.bodyText` (Task 4), bestehende `groupTasks`, `todayIso`, `makeRow`, `projectNames`.
- Produces: fertiges Nutzerverhalten; keine neuen Exports.

- [ ] **Step 1: Imports und State umstellen**

Oben in `src/taskpane/taskpane.ts`:

```ts
import { getAllTasks, getProjects, createTask, deleteComment, TodoistTask } from "../lib/todoist";
import { groupTasks, priorityColor, taskDeepLink, todayIso, filterTasks, dueTodayOrOverdue, extractMailKeywords, suggestTasks } from "./taskLogic";
```

Neuer Modul-State neben `prepared`:

```ts
let allTasks: TodoistTask[] = [];
```

- [ ] **Step 2: `renderTasks` durch `renderSections` + `renderDefaultView` ersetzen**

`renderTasks` komplett ersetzen durch:

```ts
function renderSections(sections: Array<[string, TodoistTask[]]>, emptyText: string): void {
  setSkeleton(false);
  const groups = $("task-groups");
  groups.innerHTML = "";
  const empty = $("empty");

  const total = sections.reduce((n, [, list]) => n + list.length, 0);
  if (total === 0) { empty.textContent = emptyText; empty.hidden = false; return; }
  empty.hidden = true;

  for (const [label, list] of sections) {
    if (list.length === 0) continue;
    const h = document.createElement("p");
    h.className = "group-label";
    h.textContent = label;
    groups.appendChild(h);
    const ul = document.createElement("ul");
    ul.className = "task-list";
    for (const task of list) ul.appendChild(makeRow(task));
    groups.appendChild(ul);
  }
}

// Standardansicht (leeres Suchfeld): Vorschlaege zuoberst, dann Ueberfaellig/Heute.
// Vorgeschlagene Tasks tauchen nicht doppelt in den Datumsgruppen auf.
function renderDefaultView(): void {
  const kw = extractMailKeywords(prepared?.subject ?? "", prepared?.bodyText ?? "");
  const suggestions = suggestTasks(allTasks, kw);
  const suggestedIds = new Set(suggestions.map((t) => t.id));
  const today = todayIso(new Date());
  const base = dueTodayOrOverdue(allTasks, today).filter((t) => !suggestedIds.has(t.id));
  const { overdue, today: todayList } = groupTasks(base, today);
  renderSections(
    [["Vorschläge", suggestions], ["Überfällig", overdue], ["Heute", todayList]],
    "Keine Tasks für heute.",
  );
}
```

- [ ] **Step 3: `loadTasks` auf Voll-Load umstellen**

```ts
async function loadTasks(token: string): Promise<void> {
  setSkeleton(true);
  try {
    const [tasks] = await Promise.all([
      getAllTasks(token),
      loadProjects(token),
    ]);
    allTasks = tasks;
    renderDefaultView();
  } catch (e) {
    setSkeleton(false);
    setStatus(`Token ungültig oder Abruf fehlgeschlagen: ${(e as Error).message}`, "err", e);
    showTokenSection();
  }
}
```

- [ ] **Step 4: `wireSearch` auf Client-Suche umstellen (Debounce + API-Call entfallen)**

```ts
function wireSearch(): void {
  ($("search") as HTMLInputElement).addEventListener("input", (e) => {
    const q = (e.target as HTMLInputElement).value.trim();
    if (!q) { renderDefaultView(); return; }
    renderSections([["Treffer", filterTasks(allTasks, q, projectNames)]], "Keine Treffer.");
  });

  ($("search") as HTMLInputElement).addEventListener("keydown", (e) => {
    if ((e as KeyboardEvent).key !== "Enter") return;
    const first = $("task-groups").querySelector(".task-row") as HTMLButtonElement | null;
    if (first && !first.disabled) first.click();
  });
}
```

(Der Enter-Handler bleibt identisch; in der Standardansicht ist der oberste Treffer jetzt der beste Vorschlag.)

- [ ] **Step 5: Build + alle Tests**

Run: `npx jest && npm run build`
Expected: alle Tests PASS, Webpack-Build clean (keine TS-Fehler; `renderTasks`/`searchTasks`/`getTasks` werden in taskpane.ts nirgends mehr referenziert).

- [ ] **Step 6: Commit**

```bash
git add src/taskpane/taskpane.ts
git commit -m "Task-Pane: Voll-Load aller Tasks, Client-Suche case-insensitiv, Vorschläge-Sektion"
```

---

### Task 6: Tote Server-Suche entfernen + Doku

**Files:**
- Modify: `src/lib/todoist.ts` (entfernen: `tasksByQuery`, `getTasks`, `searchTasks`)
- Modify: `tests/todoist.test.ts` (deren Tests entfernen)
- Modify: `CLAUDE.md` (Modul-Beschreibung + Status)

**Interfaces:**
- Consumes: nichts Neues.
- Produces: `todoist.ts` ohne `getTasks`/`searchTasks`/`tasksByQuery`. Alle übrigen Exports unverändert.

- [ ] **Step 1: Funktionen und Tests entfernen**

In `src/lib/todoist.ts` die Funktionen `tasksByQuery`, `getTasks`, `searchTasks` löschen. In `tests/todoist.test.ts` die `describe("getTasks")`-Suite löschen und die Imports bereinigen (der "nacktes Array"-Defensivfall und der 401-Fall sind durch die `getAllTasks`-Suite abgedeckt; falls der Array-Fall dort fehlt, als Test ergänzen):

```ts
test("kommt auch mit nacktem Array klar (defensiv)", async () => {
  (global as any).fetch = jest.fn().mockResolvedValue({
    ok: true, status: 200, json: async () => [{ id: "1", content: "Test", project_id: "p1" }], text: async () => "",
  });
  const tasks = await getAllTasks("tok");
  expect(tasks).toHaveLength(1);
});
```

- [ ] **Step 2: Alle Tests + Build**

Run: `npx jest && npm run build`
Expected: PASS + Build clean. Grep-Check: `grep -rn "getTasks\|searchTasks" src/` liefert keine Treffer mehr.

- [ ] **Step 3: CLAUDE.md nachführen**

In `CLAUDE.md`: Modulzeile `src/lib/todoist.ts` auf `getAllTasks` (Cursor-Pagination) umformulieren, `src/taskpane/taskLogic.ts` um `filterTasks/dueTodayOrOverdue/extractMailKeywords/suggestTasks` ergänzen, Taskpane-Zeile um "Vorschläge-Sektion + Client-Suche" ergänzen, neuen Status-Eintrag mit Datum 2026-07-06 anlegen (Was: case-insensitive Client-Suche + Top-3-Vorschläge, Warum: Todoist-Server-Suche matchte case-sensitiv, Testzahl aktualisieren).

- [ ] **Step 4: Commit**

```bash
git add src/lib/todoist.ts tests/todoist.test.ts CLAUDE.md
git commit -m "Cleanup: server-seitige Task-Suche entfernt, Doku auf Client-Suche-Stand"
```

---

## Verifikation nach Abschluss

- `npx jest`: alle Tests grün (39 bestehende + ~16 neue).
- `npm run build`: clean.
- Manuell (Manuel, on-device): Pane öffnen, (1) "sap" tippen findet "SAP..."-Tasks, (2) Vorschläge-Sektion erscheint bei passender Mail zuoberst, (3) Enter hängt obersten Treffer an, (4) Anhängen + Rückgängig funktionieren unverändert.
- Deploy (`npm run deploy`) erst nach Manuels Okay; vorher `gh auth status` prüfen (Account `manuelweingartner`).
