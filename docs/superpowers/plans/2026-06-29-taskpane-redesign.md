# Task-Pane-Redesign Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Die Outlook-Seitenleiste „Mail an Todoist" im Todoist-Look neu gestalten (Branding, Dark-Mode, fuenf saubere Zustaende) und neun UX-Features ergaenzen, ohne das Manifest anzufassen.

**Architecture:** Reine Logik (Todoist-API, .eml-Vorbereitung, Task-Gruppierung) bleibt in testbaren Funktionen unter `src/lib/` und neu `src/taskpane/taskLogic.ts`. `taskpane.ts` macht nur noch DOM-Verdrahtung der Zustaende. CSS wird self-contained (Fluent-CDN raus), Farben als CSS-Variablen mit Light-/Dark-Set.

**Tech Stack:** TypeScript, Office.js, Webpack, Jest (ts-jest, jsdom). Kein neues Dependency.

## Global Constraints

- **Kein Manifest-Eingriff.** Nur gehostetes UI aendern; Deploy via `npm run deploy` (Self-Deploy, kein IT-Rollout). `manifest.prod.xml` bleibt unveraendert.
- **Keine Em-/En-Dashes** (`—` / `–`) irgendwo (Code, Kommentar, UI-Text). Ersatz: `.` `,` `:` `-` `(...)`.
- **Echte Umlaute** in user-facing Deutsch (ae/oe/ue nur in Code-Bezeichnern/Kommentaren wo schon ueblich; UI-Text mit ä/ö/ü/ß).
- **Kein stilles try/catch:** Fehler immer `console.error(msg, cause)` UND sichtbar im Pane.
- **TDD:** Neue Logik kommt mit Test zuerst. Test-Runner: `npx jest`.
- **Git-Identitaet:** Manuel Weingartner <manuel.weingartner@gmx.ch>. NIE CMI. NIE Co-Authored-By Claude/Anthropic.
- **Akzentfarbe** Todoist-Rot `#E44332`.

---

### Task 1: Todoist-API erweitern (Daten + neue Endpunkte)

**Files:**
- Modify: `src/lib/todoist.ts`
- Test: `tests/todoist.test.ts`

**Interfaces:**
- Consumes: vorhandene Helfer `auth()`, `ensureOk()`, `unwrap()`, `TodoistError`, Konstante `API`.
- Produces:
  - `interface TodoistTask { id: string; content: string; project_id: string; priority?: number; due?: { date: string; datetime?: string } | null; }`
  - `interface TodoistProject { id: string; name: string; }`
  - `addComment(token, taskId, file, content?) : Promise<string>` (gibt jetzt die neue Kommentar-id zurueck)
  - `deleteComment(token: string, commentId: string): Promise<void>`
  - `getProjects(token: string): Promise<TodoistProject[]>`
  - `createTask(token: string, content: string): Promise<TodoistTask>`

- [ ] **Step 1: Tests schreiben (failing)**

In `tests/todoist.test.ts` ans Ende anfuegen:

```ts
import {
  getTasks as _gt, // (bestehender Import bleibt; nur ergaenzen falls noetig)
} from "../src/lib/todoist";
import { deleteComment, getProjects, createTask } from "../src/lib/todoist";

describe("addComment gibt id zurueck", () => {
  test("liefert die neue Kommentar-id", async () => {
    mockFetch(200, { id: "c99" });
    const id = await addComment("tok", "t1", { file_url: "u", file_name: "m.eml", file_type: "message/rfc822" }, "Betreff (01.01.2026)");
    expect(id).toBe("c99");
    const [, opts] = (global as any).fetch.mock.calls[0];
    expect(JSON.parse(opts.body).content).toBe("Betreff (01.01.2026)");
  });
});

describe("deleteComment", () => {
  test("DELETE auf /comments/:id mit Bearer", async () => {
    mockFetch(204, {});
    await deleteComment("tok", "c99");
    const [url, opts] = (global as any).fetch.mock.calls[0];
    expect(url).toBe("https://api.todoist.com/api/v1/comments/c99");
    expect(opts.method).toBe("DELETE");
    expect(opts.headers.Authorization).toBe("Bearer tok");
  });
  test("wirft TodoistError bei 404", async () => {
    mockFetch(404, { error: "not found" });
    await expect(deleteComment("tok", "x")).rejects.toBeInstanceOf(TodoistError);
  });
});

describe("getProjects", () => {
  test("liest /projects, unwrappt results", async () => {
    mockFetch(200, { results: [{ id: "p1", name: "Inbox" }], next_cursor: null });
    const ps = await getProjects("tok");
    expect(ps[0].name).toBe("Inbox");
    const [url] = (global as any).fetch.mock.calls[0];
    expect(url).toBe("https://api.todoist.com/api/v1/projects");
  });
});

describe("createTask", () => {
  test("POSTet content an /tasks und gibt Task zurueck", async () => {
    mockFetch(200, { id: "t5", content: "Neuer Task", project_id: "p1" });
    const t = await createTask("tok", "Neuer Task");
    expect(t.id).toBe("t5");
    const [url, opts] = (global as any).fetch.mock.calls[0];
    expect(url).toBe("https://api.todoist.com/api/v1/tasks");
    expect(opts.method).toBe("POST");
    expect(JSON.parse(opts.body).content).toBe("Neuer Task");
  });
});
```

Stelle sicher, dass `addComment` und `TodoistError` oben in der Datei importiert sind (sie sind es bereits).

- [ ] **Step 2: Tests laufen lassen (muessen fehlschlagen)**

Run: `npx jest tests/todoist.test.ts`
Expected: FAIL (`deleteComment`/`getProjects`/`createTask` not exported; `addComment` returns undefined).

- [ ] **Step 3: Implementierung in `src/lib/todoist.ts`**

`TodoistTask` ersetzen und `TodoistProject` ergaenzen:

```ts
export interface TodoistTask {
  id: string;
  content: string;
  project_id: string;
  priority?: number;                                  // 1 bis 4, 4 = hoechste (P1)
  due?: { date: string; datetime?: string } | null;
}
export interface TodoistProject { id: string; name: string; }
```

`addComment` so aendern, dass es die id zurueckgibt:

```ts
export async function addComment(token: string, taskId: string, file: UploadedFile, content = ""): Promise<string> {
  const res = await fetch(`${API}/comments`, {
    method: "POST",
    headers: { ...auth(token), "Content-Type": "application/json" },
    body: JSON.stringify({
      task_id: taskId,
      content: content || file.file_name,
      attachment: { resource_type: "file", file_url: file.file_url, file_name: file.file_name, file_type: file.file_type },
    }),
  });
  const data = await (await ensureOk(res)).json();
  return (data && data.id) as string;
}
```

Neue Funktionen ans Dateiende:

```ts
export async function deleteComment(token: string, commentId: string): Promise<void> {
  const res = await fetch(`${API}/comments/${commentId}`, { method: "DELETE", headers: auth(token) });
  await ensureOk(res);
}

export async function getProjects(token: string): Promise<TodoistProject[]> {
  const res = await fetch(`${API}/projects`, { headers: auth(token) });
  return unwrap<TodoistProject>(await (await ensureOk(res)).json());
}

export async function createTask(token: string, content: string): Promise<TodoistTask> {
  const res = await fetch(`${API}/tasks`, {
    method: "POST",
    headers: { ...auth(token), "Content-Type": "application/json" },
    body: JSON.stringify({ content }),
  });
  return (await ensureOk(res)).json();
}
```

- [ ] **Step 4: Tests laufen lassen (muessen gruen sein)**

Run: `npx jest tests/todoist.test.ts`
Expected: PASS (alle, inkl. der bestehenden).

- [ ] **Step 5: Commit**

```bash
git add src/lib/todoist.ts tests/todoist.test.ts
git commit -m "Todoist-API: priority/due, deleteComment, getProjects, createTask, addComment gibt id"
```

---

### Task 2: Mail-Vorbereitung + Anhaenge-Flow (Groessen-Vorabcheck, Betreff als Kommentar, Undo-faehige id)

**Files:**
- Modify: `src/lib/attachToTask.ts`
- Test: `tests/attachToTask.orchestration.test.ts`, neu `tests/attachToTask.prepare.test.ts`

**Interfaces:**
- Consumes: `readCurrentMail()`, `emlBlobFor()`, `uploadFile()`, `addComment()` (Task 1, gibt id), `MAX_BYTES`.
- Produces:
  - `interface PreparedMail { blob: Blob; fileName: string; sizeBytes: number; subject: string; commentText: string; }`
  - `formatMailDate(utc: string): string` (gibt `dd.mm.yyyy` oder `""`)
  - `prepareCurrentMail(): Promise<PreparedMail>`
  - `attachPreparedToTask(token: string, taskId: string, prepared: PreparedMail): Promise<string>` (gibt Kommentar-id zurueck)
  - `attachCurrentMailToTask(token, taskId): Promise<string>` (bleibt als Convenience-Wrapper)

- [ ] **Step 1: Tests schreiben (failing)**

Neue Datei `tests/attachToTask.prepare.test.ts`:

```ts
import { formatMailDate } from "../src/lib/attachToTask";

describe("formatMailDate", () => {
  test("formatiert UTC-String zu dd.mm.yyyy", () => {
    expect(formatMailDate("Mon, 05 Jan 2026 10:00:00 +0000")).toBe("05.01.2026");
  });
  test("leerer/ungueltiger String gibt leer", () => {
    expect(formatMailDate("")).toBe("");
    expect(formatMailDate("kein datum")).toBe("");
  });
});
```

`tests/attachToTask.orchestration.test.ts` anpassen: den happy-path-Test so aendern, dass der Kommentartext geprueft wird und die id zurueckkommt. Ersetze die `addComment`-Assertion und ergaenze einen Prepared-Test:

```ts
import { attachCurrentMailToTask, attachPreparedToTask, prepareCurrentMail, MAX_BYTES } from "../src/lib/attachToTask";
// ... baseMail wie gehabt ...

test("happy path: Kommentartext = Betreff (Datum), gibt id zurueck", async () => {
  (readCurrentMail as jest.Mock).mockResolvedValue(baseMail);
  (uploadFile as jest.Mock).mockResolvedValue({ file_url: "u", file_name: "S.eml", file_type: "message/rfc822" });
  (addComment as jest.Mock).mockResolvedValue("c1");

  const id = await attachCurrentMailToTask("tok", "task1");

  expect(id).toBe("c1");
  const [tok, taskId, file, content] = (addComment as jest.Mock).mock.calls[0];
  expect(tok).toBe("tok");
  expect(taskId).toBe("task1");
  expect(file.file_url).toBe("u");
  expect(content).toBe("S (01.01.2026)");
});

test("prepareCurrentMail liefert sizeBytes + commentText", async () => {
  (readCurrentMail as jest.Mock).mockResolvedValue(baseMail);
  const p = await prepareCurrentMail();
  expect(p.subject).toBe("S");
  expect(p.commentText).toBe("S (01.01.2026)");
  expect(p.sizeBytes).toBeGreaterThan(0);
});
```

(Der bestehende ">25 MB"-Test bleibt; er ruft weiterhin `attachCurrentMailToTask` und erwartet Throw + kein Upload.)

- [ ] **Step 2: Tests laufen lassen (muessen fehlschlagen)**

Run: `npx jest tests/attachToTask.prepare.test.ts tests/attachToTask.orchestration.test.ts`
Expected: FAIL (`formatMailDate`/`prepareCurrentMail`/`attachPreparedToTask` not exported; addComment-4-Arg-Aufruf fehlt).

- [ ] **Step 3: Implementierung in `src/lib/attachToTask.ts`**

`addComment` muss bereits importiert sein. Datei so umbauen (die bestehenden `MAX_BYTES`, `sanitizeFileName`, `emlBlobFor` bleiben):

```ts
import { buildEml, MailData } from "./emlBuilder";
import { readCurrentMail } from "./mailReader";
import { uploadFile, addComment } from "./todoist";

export const MAX_BYTES = 25 * 1024 * 1024;

export interface PreparedMail {
  blob: Blob;
  fileName: string;
  sizeBytes: number;
  subject: string;
  commentText: string;
}

function sanitizeFileName(subject: string): string {
  const cleaned = (subject || "Mail").replace(/[\\/:*?"<>|]/g, "_").trim();
  return `${cleaned.slice(0, 120) || "Mail"}.eml`;
}

export function emlBlobFor(mail: MailData): { blob: Blob; fileName: string } {
  const eml = buildEml(mail);
  const blob = new Blob([eml], { type: "message/rfc822" });
  return { blob, fileName: sanitizeFileName(mail.subject) };
}

export function formatMailDate(utc: string): string {
  const d = new Date(utc);
  if (isNaN(d.getTime())) return "";
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  return `${dd}.${mm}.${d.getFullYear()}`;
}

export async function prepareCurrentMail(): Promise<PreparedMail> {
  const mail = await readCurrentMail();
  const { blob, fileName } = emlBlobFor(mail);
  const datePart = formatMailDate(mail.date);
  const commentText = datePart ? `${mail.subject} (${datePart})` : mail.subject;
  return { blob, fileName, sizeBytes: blob.size, subject: mail.subject, commentText };
}

export async function attachPreparedToTask(token: string, taskId: string, prepared: PreparedMail): Promise<string> {
  if (prepared.sizeBytes > MAX_BYTES) {
    throw new Error(`Mail ist ${(prepared.sizeBytes / 1024 / 1024).toFixed(1)} MB gross, Todoist erlaubt max 25 MB.`);
  }
  const uploaded = await uploadFile(token, prepared.blob, prepared.fileName);
  return addComment(token, taskId, uploaded, prepared.commentText);
}

export async function attachCurrentMailToTask(token: string, taskId: string): Promise<string> {
  const prepared = await prepareCurrentMail();
  return attachPreparedToTask(token, taskId, prepared);
}
```

- [ ] **Step 4: Tests laufen lassen (muessen gruen sein)**

Run: `npx jest tests/attachToTask.prepare.test.ts tests/attachToTask.orchestration.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/lib/attachToTask.ts tests/attachToTask.prepare.test.ts tests/attachToTask.orchestration.test.ts
git commit -m "Mail-Vorbereitung: prepareCurrentMail, Vorab-Groessencheck, Betreff(Datum) als Kommentar, id-Rueckgabe"
```

---

### Task 3: Reine Task-Pane-Logik (Gruppierung, Prioritaetsfarbe, Deeplink)

**Files:**
- Create: `src/taskpane/taskLogic.ts`
- Test: `tests/taskLogic.test.ts`

**Interfaces:**
- Consumes: `TodoistTask` (Task 1).
- Produces:
  - `todayIso(now: Date): string` (`YYYY-MM-DD`)
  - `groupTasks(tasks: TodoistTask[], todayIso: string): { overdue: TodoistTask[]; today: TodoistTask[] }`
  - `priorityColor(priority?: number): string`
  - `taskDeepLink(id: string): string`

- [ ] **Step 1: Test schreiben (failing)**

Neue Datei `tests/taskLogic.test.ts`:

```ts
import { todayIso, groupTasks, priorityColor, taskDeepLink } from "../src/taskpane/taskLogic";
import { TodoistTask } from "../src/lib/todoist";

const t = (id: string, due: string | null, priority = 1): TodoistTask => ({
  id, content: id, project_id: "p1", priority, due: due ? { date: due } : null,
});

describe("todayIso", () => {
  test("formatiert YYYY-MM-DD", () => {
    expect(todayIso(new Date("2026-06-29T12:00:00Z"))).toBe("2026-06-29");
  });
});

describe("groupTasks", () => {
  test("trennt ueberfaellig (vor heute) von heute/ohne Datum", () => {
    const g = groupTasks([t("a", "2026-06-28"), t("b", "2026-06-29"), t("c", null)], "2026-06-29");
    expect(g.overdue.map((x) => x.id)).toEqual(["a"]);
    expect(g.today.map((x) => x.id)).toEqual(["b", "c"]);
  });
});

describe("priorityColor", () => {
  test("P1..P4 Mapping", () => {
    expect(priorityColor(4)).toBe("#e44332");
    expect(priorityColor(3)).toBe("#eb8909");
    expect(priorityColor(2)).toBe("#246fe0");
    expect(priorityColor(1)).toBe("#808080");
    expect(priorityColor(undefined)).toBe("#808080");
  });
});

describe("taskDeepLink", () => {
  test("baut Todoist-Task-URL", () => {
    expect(taskDeepLink("t5")).toBe("https://app.todoist.com/app/task/t5");
  });
});
```

- [ ] **Step 2: Test laufen lassen (muss fehlschlagen)**

Run: `npx jest tests/taskLogic.test.ts`
Expected: FAIL (module not found).

- [ ] **Step 3: Implementierung `src/taskpane/taskLogic.ts`**

```ts
import { TodoistTask } from "../lib/todoist";

export function todayIso(now: Date): string {
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  return `${now.getFullYear()}-${mm}-${dd}`;
}

export interface GroupedTasks { overdue: TodoistTask[]; today: TodoistTask[]; }

export function groupTasks(tasks: TodoistTask[], today: string): GroupedTasks {
  const overdue: TodoistTask[] = [];
  const todayList: TodoistTask[] = [];
  for (const task of tasks) {
    const date = task.due?.date;
    if (date && date < today) overdue.push(task);
    else todayList.push(task);
  }
  return { overdue, today: todayList };
}

export function priorityColor(priority?: number): string {
  switch (priority) {
    case 4: return "#e44332"; // P1 rot
    case 3: return "#eb8909"; // P2 orange
    case 2: return "#246fe0"; // P3 blau
    default: return "#808080"; // P4 / keine
  }
}

export function taskDeepLink(id: string): string {
  return `https://app.todoist.com/app/task/${id}`;
}
```

- [ ] **Step 4: Test laufen lassen (muss gruen sein)**

Run: `npx jest tests/taskLogic.test.ts`
Expected: PASS.

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/taskLogic.ts tests/taskLogic.test.ts
git commit -m "Task-Pane-Logik: groupTasks, priorityColor, taskDeepLink, todayIso (rein, getestet)"
```

---

### Task 4: HTML-Struktur + Todoist-CSS (Branding, Dark-Mode, fuenf Zustaende, Fluent-CDN raus)

**Files:**
- Modify: `src/taskpane/taskpane.html`
- Modify: `src/taskpane/taskpane.css` (komplett ersetzen)

**Interfaces:**
- Produces (DOM-Vertrag fuer Task 5, exakte Element-ids):
  `app-header`, `context-bar`, `context-subject`, `token-section`, `token-input`, `token-save`,
  `task-section`, `search`, `list-region`, `skeleton`, `empty`, `task-groups`, `new-task`, `size-warning`, `status`.

- [ ] **Step 1: `taskpane.html` ersetzen**

```html
<!DOCTYPE html>
<html lang="de">
<head>
  <meta charset="UTF-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Mail an Todoist</title>
  <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  <link href="taskpane.css" rel="stylesheet" type="text/css" />
</head>
<body>
  <main id="app">
    <header id="app-header">
      <span class="brand-dot" aria-hidden="true"></span>
      <h1>Mail an Todoist</h1>
    </header>

    <p id="context-bar" hidden>Anhaengen: <strong id="context-subject"></strong></p>
    <p id="size-warning" hidden></p>

    <section id="token-section" hidden>
      <p class="lead">Verbinde dein Todoist-Konto, um Mails an Tasks zu haengen.</p>
      <ol class="hint">
        <li>Todoist oeffnen: Einstellungen &gt; Integrationen &gt; Entwickler</li>
        <li>API-Token kopieren und unten einfuegen</li>
      </ol>
      <input id="token-input" type="password" placeholder="API-Token" autocomplete="off" />
      <button id="token-save" class="primary">Speichern</button>
    </section>

    <section id="task-section" hidden>
      <div class="search-wrap">
        <span class="search-icon" aria-hidden="true">&#128269;</span>
        <input id="search" placeholder="Task suchen..." autocomplete="off" />
      </div>
      <div id="list-region">
        <div id="skeleton" hidden>
          <span class="skel-row"></span><span class="skel-row"></span><span class="skel-row"></span>
        </div>
        <p id="empty" hidden>Keine Tasks fuer heute.</p>
        <div id="task-groups"></div>
      </div>
      <button id="new-task" class="ghost" hidden>+ Neuen Task mit dieser Mail</button>
    </section>

    <p id="status" role="status"></p>
  </main>
</body>
</html>
```

- [ ] **Step 2: `taskpane.css` komplett ersetzen**

```css
:root {
  --accent: #e44332;
  --accent-hover: #c93b2c;
  --bg: #ffffff;
  --bg-soft: #f7f7f7;
  --bg-hover: #f0f0f0;
  --text: #202020;
  --text-soft: #707070;
  --border: #e0e0e0;
  --ok: #058527;
  --err: #b00020;
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg: #1f1f1f; --bg-soft: #2a2a2a; --bg-hover: #333333;
    --text: #f0f0f0; --text-soft: #a0a0a0; --border: #3a3a3a;
    --ok: #4ec96f; --err: #ff6b6b;
  }
}

* { box-sizing: border-box; }
html, body { width: 100%; height: 100%; margin: 0; padding: 0; }
body {
  font-family: "Segoe UI", system-ui, -apple-system, sans-serif;
  font-size: 14px; color: var(--text); background: var(--bg);
}
#app { padding: 12px; }

#app-header { display: flex; align-items: center; gap: 8px; margin-bottom: 12px; }
#app-header h1 { font-size: 15px; font-weight: 600; margin: 0; }
.brand-dot { width: 12px; height: 12px; border-radius: 50%; background: var(--accent); flex: 0 0 auto; }

#context-bar { font-size: 12px; color: var(--text-soft); margin: 0 0 10px; padding: 6px 8px; background: var(--bg-soft); border-radius: 6px; }
#context-bar strong { color: var(--text); }
#size-warning { font-size: 12px; color: var(--err); margin: 0 0 10px; padding: 6px 8px; border: 1px solid var(--err); border-radius: 6px; }

.lead { margin: 0 0 8px; }
.hint { font-size: 12px; color: var(--text-soft); margin: 0 0 12px; padding-left: 18px; }
.hint li { margin-bottom: 4px; }

input { width: 100%; padding: 8px 10px; border: 1px solid var(--border); border-radius: 6px; background: var(--bg); color: var(--text); font-size: 14px; }
input:focus { outline: none; border-color: var(--accent); }

button { font-family: inherit; font-size: 14px; cursor: pointer; border-radius: 6px; }
button.primary { width: 100%; margin-top: 8px; padding: 9px; border: none; background: var(--accent); color: #fff; font-weight: 600; }
button.primary:hover { background: var(--accent-hover); }
button.ghost { width: 100%; margin-top: 10px; padding: 8px; border: 1px dashed var(--border); background: transparent; color: var(--text-soft); }
button.ghost:hover { border-color: var(--accent); color: var(--accent); }

.search-wrap { position: relative; margin-bottom: 10px; }
.search-wrap .search-icon { position: absolute; left: 9px; top: 50%; transform: translateY(-50%); font-size: 13px; opacity: 0.6; }
.search-wrap input { padding-left: 30px; }

.group-label { font-size: 11px; font-weight: 700; letter-spacing: 0.04em; text-transform: uppercase; color: var(--text-soft); margin: 12px 0 4px; }

ul.task-list { list-style: none; margin: 0; padding: 0; }
.task-row { display: flex; align-items: center; gap: 8px; width: 100%; text-align: left; padding: 9px 8px; margin: 2px 0; border: none; background: transparent; color: var(--text); border-radius: 6px; }
.task-row:hover { background: var(--bg-hover); }
.task-row:focus-visible { outline: 2px solid var(--accent); }
.task-row .prio { width: 8px; height: 8px; border-radius: 50%; flex: 0 0 auto; }
.task-row .content { flex: 1 1 auto; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.task-row .project { font-size: 11px; color: var(--text-soft); flex: 0 0 auto; }
.task-row .state { font-size: 12px; flex: 0 0 auto; }
.task-row .state.ok { color: var(--ok); }
.task-row .state.busy { color: var(--text-soft); }

.row-open { background: none; border: none; padding: 2px 4px; color: var(--text-soft); cursor: pointer; font-size: 12px; }
.row-open:hover { color: var(--accent); }

.undo { background: none; border: none; color: var(--accent); cursor: pointer; font-size: 12px; padding: 0; text-decoration: underline; }

#skeleton { display: flex; flex-direction: column; gap: 8px; padding: 8px 0; }
.skel-row { height: 16px; border-radius: 4px; background: linear-gradient(90deg, var(--bg-soft) 25%, var(--bg-hover) 50%, var(--bg-soft) 75%); background-size: 200% 100%; animation: skel 1.2s infinite; }
@keyframes skel { 0% { background-position: 200% 0; } 100% { background-position: -200% 0; } }

#empty { color: var(--text-soft); text-align: center; padding: 20px 0; }
#status { margin-top: 10px; font-size: 13px; min-height: 18px; }
#status.err { color: var(--err); }
#status.ok { color: var(--ok); }
```

- [ ] **Step 3: Build pruefen**

Run: `npm run build`
Expected: Webpack-Build ohne Fehler, `dist/taskpane.html` + `dist/taskpane.css` erzeugt. (Verdrahtung folgt in Task 5; UI ist hier noch statisch.)

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.html src/taskpane/taskpane.css
git commit -m "Task-Pane: Todoist-Branding, Dark-Mode, fuenf-Zustaende-Markup, Fluent-CDN entfernt"
```

---

### Task 5: `taskpane.ts` neu verdrahten (Zustaende, Inline-Feedback, Undo, Tastatur, Groessencheck, Projektnamen, neuer Task, Open-in-Todoist)

**Files:**
- Modify: `src/taskpane/taskpane.ts` (komplett ersetzen)

**Interfaces:**
- Consumes: alles aus Tasks 1-4 (`getTasks`, `searchTasks`, `getProjects`, `createTask`, `deleteComment`, `TodoistTask`, `TodoistProject`; `prepareCurrentMail`, `attachPreparedToTask`, `PreparedMail`; `groupTasks`, `priorityColor`, `taskDeepLink`, `todayIso`; `getToken`, `setToken`).
- Produces: kein Export (Entry-Point).

- [ ] **Step 1: `src/taskpane/taskpane.ts` komplett ersetzen**

```ts
import { getToken, setToken } from "../lib/settings";
import { getTasks, searchTasks, getProjects, createTask, deleteComment, TodoistTask } from "../lib/todoist";
import { prepareCurrentMail, attachPreparedToTask, PreparedMail } from "../lib/attachToTask";
import { groupTasks, priorityColor, taskDeepLink, todayIso } from "./taskLogic";

let busy = false;
let searchWired = false;
let prepared: PreparedMail | null = null;
let projectNames: Record<string, string> = {};

function $(id: string): HTMLElement { return document.getElementById(id)!; }

function setStatus(msg: string, kind: "" | "ok" | "err" = "", cause?: unknown): void {
  const el = $("status");
  el.textContent = msg;
  el.className = kind;
  if (kind === "err") { if (cause !== undefined) console.error(msg, cause); else console.error(msg); }
}

function showTokenSection(): void { $("token-section").hidden = false; $("task-section").hidden = true; }
function showTaskSection(): void { $("token-section").hidden = true; $("task-section").hidden = false; }

function renderContextBar(): void {
  if (!prepared) return;
  $("context-subject").textContent = prepared.subject || "(kein Betreff)";
  $("context-bar").hidden = false;
}

function renderSizeWarning(): void {
  const MAX = 25 * 1024 * 1024;
  const warn = $("size-warning");
  if (prepared && prepared.sizeBytes > MAX) {
    warn.textContent = `Mail ist ${(prepared.sizeBytes / 1024 / 1024).toFixed(1)} MB gross. Todoist erlaubt max 25 MB, Anhaengen ist deaktiviert.`;
    warn.hidden = false;
  } else {
    warn.hidden = true;
  }
}

function tooLarge(): boolean {
  return !!prepared && prepared.sizeBytes > 25 * 1024 * 1024;
}

function setSkeleton(on: boolean): void { $("skeleton").hidden = !on; }

function makeRow(task: TodoistTask): HTMLLIElement {
  const li = document.createElement("li");

  const btn = document.createElement("button");
  btn.className = "task-row";
  btn.disabled = tooLarge();

  const dot = document.createElement("span");
  dot.className = "prio";
  dot.style.background = priorityColor(task.priority);
  btn.appendChild(dot);

  const content = document.createElement("span");
  content.className = "content";
  content.textContent = task.content;
  btn.appendChild(content);

  const projectName = projectNames[task.project_id];
  if (projectName) {
    const proj = document.createElement("span");
    proj.className = "project";
    proj.textContent = projectName;
    btn.appendChild(proj);
  }

  const state = document.createElement("span");
  state.className = "state";
  btn.appendChild(state);

  btn.onclick = () => attach(task, btn, state, li);

  const open = document.createElement("button");
  open.className = "row-open";
  open.title = "In Todoist oeffnen";
  open.textContent = "↗"; // Pfeil nach oben rechts
  open.onclick = (e) => { e.stopPropagation(); window.open(taskDeepLink(task.id), "_blank", "noopener"); };

  li.appendChild(btn);
  li.appendChild(open);
  return li;
}

function renderTasks(tasks: TodoistTask[]): void {
  setSkeleton(false);
  const groups = $("task-groups");
  groups.innerHTML = "";
  const empty = $("empty");

  if (tasks.length === 0) { empty.hidden = false; return; }
  empty.hidden = true;

  const { overdue, today } = groupTasks(tasks, todayIso(new Date()));
  const sections: Array<[string, TodoistTask[]]> = [["Ueberfaellig", overdue], ["Heute", today]];
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

async function attach(task: TodoistTask, btn: HTMLButtonElement, state: HTMLElement, li: HTMLElement): Promise<void> {
  if (busy || tooLarge() || !prepared) return;
  const token = getToken();
  if (!token) { showTokenSection(); return; }
  busy = true;
  btn.disabled = true;
  state.className = "state busy";
  state.textContent = "...";
  try {
    const commentId = await attachPreparedToTask(token, task.id, prepared);
    state.className = "state ok";
    state.textContent = "✓"; // Haken
    renderUndo(li, token, commentId, state);
    setStatus("", "");
  } catch (e) {
    state.className = "state";
    state.textContent = "";
    btn.disabled = false;
    setStatus(`Fehler: ${(e as Error).message}`, "err", e);
  } finally {
    busy = false;
  }
}

function renderUndo(li: HTMLElement, token: string, commentId: string, state: HTMLElement): void {
  if (!commentId) return;
  const undo = document.createElement("button");
  undo.className = "undo";
  undo.textContent = "Rueckgaengig";
  undo.onclick = async () => {
    undo.disabled = true;
    try {
      await deleteComment(token, commentId);
      state.className = "state";
      state.textContent = "";
      undo.remove();
      setStatus("Anhang entfernt.", "ok");
    } catch (e) {
      undo.disabled = false;
      setStatus(`Rueckgaengig fehlgeschlagen: ${(e as Error).message}`, "err", e);
    }
  };
  li.appendChild(undo);
}

async function loadTasks(token: string): Promise<void> {
  setSkeleton(true);
  try {
    const [tasks] = await Promise.all([
      getTasks(token),
      loadProjects(token),
    ]);
    renderTasks(tasks);
  } catch (e) {
    setSkeleton(false);
    setStatus(`Token ungueltig oder Abruf fehlgeschlagen: ${(e as Error).message}`, "err", e);
    showTokenSection();
  }
}

async function loadProjects(token: string): Promise<void> {
  try {
    const ps = await getProjects(token);
    projectNames = Object.fromEntries(ps.map((p) => [p.id, p.name]));
  } catch (e) {
    console.error("Projekte konnten nicht geladen werden", e); // nicht fatal: Tasks ohne Projektlabel
  }
}

function wireSearch(): void {
  let timer: number | undefined;
  ($("search") as HTMLInputElement).addEventListener("input", (e) => {
    const q = (e.target as HTMLInputElement).value.trim();
    const token = getToken();
    if (!token) return;
    window.clearTimeout(timer);
    timer = window.setTimeout(async () => {
      setSkeleton(true);
      try {
        const tasks = q ? await searchTasks(token, q) : await getTasks(token);
        renderTasks(tasks);
      } catch (err) {
        setSkeleton(false);
        setStatus(`Suche fehlgeschlagen: ${(err as Error).message}`, "err", err);
      }
    }, 300);
  });

  ($("search") as HTMLInputElement).addEventListener("keydown", (e) => {
    if ((e as KeyboardEvent).key !== "Enter") return;
    const first = $("task-groups").querySelector(".task-row") as HTMLButtonElement | null;
    if (first && !first.disabled) first.click();
  });
}

function wireNewTask(): void {
  const btn = $("new-task") as HTMLButtonElement;
  btn.hidden = false;
  btn.onclick = async () => {
    if (busy || tooLarge() || !prepared) return;
    const token = getToken();
    if (!token) { showTokenSection(); return; }
    busy = true; btn.disabled = true;
    setStatus("Erstelle Task...", "");
    try {
      const task = await createTask(token, prepared.subject || "Mail");
      const commentId = await attachPreparedToTask(token, task.id, prepared);
      setStatus(`Neuer Task "${task.content}" mit Mail erstellt.`, "ok");
      void commentId;
      await loadTasks(token);
    } catch (e) {
      setStatus(`Fehler beim Erstellen: ${(e as Error).message}`, "err", e);
    } finally {
      busy = false; btn.disabled = false;
    }
  };
}

async function start(): Promise<void> {
  const token = getToken();
  if (!token) { showTokenSection(); return; }
  showTaskSection();

  try {
    prepared = await prepareCurrentMail();
    renderContextBar();
    renderSizeWarning();
  } catch (e) {
    setStatus(`Mail konnte nicht gelesen werden: ${(e as Error).message}`, "err", e);
  }

  if (!searchWired) { wireSearch(); wireNewTask(); searchWired = true; }
  ($("search") as HTMLInputElement).focus();
  await loadTasks(token);
}

Office.onReady(() => {
  ($("token-save") as HTMLButtonElement).onclick = async () => {
    const val = ($("token-input") as HTMLInputElement).value.trim();
    if (!val) { setStatus("Bitte Token eingeben.", "err"); return; }
    await setToken(val);
    void start();
  };
  void start();
});
```

- [ ] **Step 2: Volle Test-Suite + Build**

Run: `npx jest`
Expected: PASS (alle, inkl. `tests/smoke.test.ts`).

Run: `npm run build`
Expected: Webpack-Build ohne Fehler.

- [ ] **Step 3: Lint**

Run: `npx eslint src --ext .ts` (falls konfiguriert; sonst ueberspringen)
Expected: keine Fehler (insbesondere keine ungenutzten Importe).

- [ ] **Step 4: Commit**

```bash
git add src/taskpane/taskpane.ts
git commit -m "Task-Pane verdrahtet: Zustaende, Inline-Feedback, Undo, Tastatur-Flow, Groessencheck, Projektnamen, neuer Task, Open-in-Todoist"
```

---

### Task 6: Verifikation am echten Konto + Deploy

**Files:** keine Code-Aenderung (Verifikations-Gate). Falls die API-Felder abweichen, zurueck zu Task 1/3.

- [ ] **Step 1: Manifest-Unveraendertheit pruefen**

Run: `git status --porcelain manifest.prod.xml`
Expected: leere Ausgabe (Manifest nicht angefasst, kein IT-Rollout noetig).

Run: `npx office-addin-manifest validate manifest.prod.xml`
Expected: „is valid".

- [ ] **Step 2: API-Felder empirisch verifizieren (Konvention: konkrete Fakten verifizieren)**

Mit echtem Token gegen die v1-API pruefen, dass `priority` und `due.date` so geliefert werden wie angenommen (sonst Mapping in Task 1/3 anpassen):

```bash
curl -s -H "Authorization: Bearer <TOKEN>" \
  "https://api.todoist.com/api/v1/tasks/filter?query=$(printf '%s' '(today | overdue)' | sed 's/ /%20/g; s/(/%28/g; s/)/%29/g; s/|/%7C/g')" | head -c 1200
```
Erwartet: Objekte mit `priority` (1..4) und `due: { date: "YYYY-MM-DD", ... }`.

Deeplink am Konto pruefen: `https://app.todoist.com/app/task/<id>` oeffnet den Task.

- [ ] **Step 3: Optik visuell abnehmen (Design-Entscheidung beim Nutzer)**

Lokal `npm run dev-server` starten und `https://localhost:3000/taskpane.html` im Browser oeffnen (Office.js-Aufrufe schlagen ausserhalb Outlook fehl, aber Layout/Light/Dark/Skeleton/Token-Screen sind sichtbar). Screenshot an Manuel; Light- und Dark-Mode zeigen. Erst nach seinem OK deployen (Praeferenz: visuelle/UX-Calls bestaetigt der Nutzer).

- [ ] **Step 4: Deploy (manuell, Account-Switch beachten)**

```bash
gh auth switch --user manuelweingartner
npm run deploy
gh auth switch --user CMI-Kunden
```
Expected: `dist/` nach `gh-pages` gepusht; Live unter https://manuelweingartner.github.io/outlook-todoist-addin/taskpane.html. Add-in in Outlook oeffnen (Cache ggf. leeren), End-to-End ein echtes Mail anhaengen + Rueckgaengig testen.

- [ ] **Step 5: CLAUDE.md-Status aktualisieren + Commit**

Im Abschnitt „Status / Gotchas" einen Eintrag mit Datum 2026-06-29 zum Redesign + den 9 Features ergaenzen.

```bash
git add CLAUDE.md
git commit -m "Doku: Task-Pane-Redesign 2026-06-29 (Branding, Dark-Mode, 9 Features) deployt"
```

---

## Self-Review

**Spec coverage:**
- Branding/Rot/Self-contained-CSS/Dark-Mode -> Task 4. ✔
- Fuenf Zustaende (Onboarding/Skeleton/Liste/Leer/Inline-Anhaengen) -> Task 4 (Markup/CSS) + Task 5 (Verhalten). ✔
- Ein-Klick-Anhaengen bleibt -> Task 5 `attach()`. ✔
- Feature 1 Undo -> Task 1 `deleteComment` + Task 5 `renderUndo`. ✔
- Feature 2 Betreff(Datum) als Kommentar -> Task 2 `prepareCurrentMail`/`commentText`. ✔
- Feature 3 Kontext-Kopf -> Task 4 `context-bar` + Task 5 `renderContextBar`. ✔
- Feature 5 (neuer Task) -> Task 1 `createTask` + Task 5 `wireNewTask`. ✔
- Feature 6 Vorab-Groessencheck -> Task 2 `sizeBytes` + Task 5 `renderSizeWarning`/`tooLarge`. ✔
- Feature 7 Skeleton -> Task 4 CSS + Task 5 `setSkeleton`. ✔
- Feature 8 Tastatur-Flow -> Task 5 `wireSearch` focus + Enter. ✔
- Feature 9 Projektname -> Task 1 `getProjects` + Task 5 `loadProjects`/`projectNames`. ✔
- Feature 10 Open-in-Todoist -> Task 3 `taskDeepLink` + Task 5 `.row-open`. ✔
- Daten `priority`/`due` -> Task 1 + Task 3 `groupTasks`/`priorityColor`. ✔
- Verifikation API-Felder + Deeplink + Manifest unveraendert -> Task 6. ✔

**Placeholder scan:** keine TBD/TODO; jeder Code-Step zeigt vollstaendigen Code.

**Type consistency:** `attachPreparedToTask`/`prepareCurrentMail`/`PreparedMail` einheitlich (Task 2 ↔ Task 5); `getProjects: Promise<TodoistProject[]>` ↔ `projectNames` Map; `addComment: Promise<string>` ↔ `commentId` in Undo. `groupTasks(tasks, today)` Signatur ↔ Aufruf mit `todayIso(new Date())`.
