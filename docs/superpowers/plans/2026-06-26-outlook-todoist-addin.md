# Outlook-Add-in „Mail an Todoist-Task" Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Ein Outlook-Web-Add-in, das die aktuell geöffnete Mail als `.eml` an einen per Klick gewählten bestehenden Todoist-Task hängt.

**Architecture:** Statisches Office.js-Task-Pane (kein Backend). Mail wird rein über Office.js gelesen und client-seitig zu einer `.eml` zusammengebaut (kein Microsoft Graph). Die `.eml` wird in die Todoist-Sync-Upload-API hochgeladen und per REST-Kommentar an den gewählten Task gehängt. Auth gegen Todoist via persönlichem API-Token (CORS erlaubt Browser-Aufrufe).

**Tech Stack:** TypeScript, Office.js (Mailbox Requirement Set 1.8), Webpack (via `yo office`-Scaffold), Jest (Unit-Tests), GitHub Pages (Hosting), klassisches XML-Add-in-Manifest.

## Global Constraints

- **Mailbox Requirement Set 1.8** ist Minimum (für `getAttachmentContentAsync`) — im Manifest setzen.
- **Kein Microsoft Graph, kein `Mail.Read`-Consent, kein Azure-App.** Nur Office.js `ReadItem`.
- **Kein eigenes Backend/Proxy.** Todoist direkt aus dem Browser (Bearer-Token, credentials-Mode `omit`).
- **Keine stillen `try/catch`:** jeder Fehler wird geloggt (`console.error(e)`) UND dem Nutzer im Pane sichtbar gemeldet.
- **Keine Em-/En-Dashes** (`—`/`–`) in Code, Kommentaren, UI-Texten. Nur `.`/`,`/`:`/`-`.
- **Git-Identität:** Manuel Weingartner <manuel.weingartner@gmx.ch>. Niemals Co-Authored-By Claude.
- Todoist-Upload-Limit: **25 MB** pro Datei.
- Todoist-Endpunkte: REST `https://api.todoist.com/rest/v2`, Sync `https://api.todoist.com/sync/v9`.

---

### Task 1: Projekt-Scaffold + Test-Setup

**Files:**
- Create: gesamtes `yo office`-Gerüst (`package.json`, `webpack.config.js`, `src/taskpane/*`, `manifest.xml`, `tsconfig.json`)
- Create: `jest.config.js`
- Create: `tests/setup.ts`
- Test: `tests/smoke.test.ts`

**Interfaces:**
- Consumes: nichts (erster Task).
- Produces: lauffähiges Build- und Test-Setup. `npx jest` läuft.

- [ ] **Step 1: Scaffold generieren (nicht-interaktiv)**

Run im Repo-Root `C:\CLAUDE\outlook-todoist-addin`:
```bash
npx --yes yo office --projectType taskpane --name "Mail an Todoist" --host outlook --ts true --skip-install false
```
Erwartet: Ordner `src/taskpane/`, `manifest.xml`, `package.json` entstehen. Falls der Generator in einen Unterordner schreibt, Inhalte ins Repo-Root verschieben.

- [ ] **Step 2: Jest + Typen installieren**

```bash
npm install --save-dev jest ts-jest @types/jest @types/office-js
```

- [ ] **Step 3: `jest.config.js` anlegen**

```js
module.exports = {
  preset: "ts-jest",
  testEnvironment: "jsdom",
  setupFilesAfterEnv: ["<rootDir>/tests/setup.ts"],
  testMatch: ["<rootDir>/tests/**/*.test.ts"],
};
```
Falls `jsdom` fehlt: `npm install --save-dev jest-environment-jsdom`.

- [ ] **Step 4: `tests/setup.ts` mit Office-Mock-Gerüst anlegen**

```ts
// Minimaler globaler Office-Mock. Einzelne Tests überschreiben Felder gezielt.
(global as any).Office = {
  CoercionType: { Html: "html", Text: "text" },
  AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
  MailboxEnums: {
    AttachmentType: { File: "file", Item: "item", Cloud: "cloud" },
    AttachmentContentFormat: { Base64: "base64", Url: "url", Eml: "eml" },
  },
  context: {},
};
```

- [ ] **Step 5: Smoke-Test schreiben**

```ts
// tests/smoke.test.ts
test("test runner laeuft", () => {
  expect(1 + 1).toBe(2);
});
```

- [ ] **Step 6: Test ausführen**

Run: `npx jest tests/smoke.test.ts`
Expected: PASS (1 test).

- [ ] **Step 7: Commit**

```bash
git add -A
git commit -m "Scaffold: Office-Taskpane-Add-in + Jest-Setup"
```

---

### Task 2: Settings-Modul (Todoist-Token-Persistenz)

**Files:**
- Create: `src/lib/settings.ts`
- Test: `tests/settings.test.ts`

**Interfaces:**
- Consumes: `Office.context.roamingSettings`.
- Produces:
  - `getToken(): string | null`
  - `setToken(token: string): Promise<void>`

- [ ] **Step 1: Failing test schreiben**

```ts
// tests/settings.test.ts
import { getToken, setToken } from "../src/lib/settings";

describe("settings", () => {
  let store: Record<string, unknown>;
  beforeEach(() => {
    store = {};
    (global as any).Office.context.roamingSettings = {
      get: (k: string) => store[k],
      set: (k: string, v: unknown) => { store[k] = v; },
      saveAsync: (cb: (r: any) => void) =>
        cb({ status: (global as any).Office.AsyncResultStatus.Succeeded }),
    };
  });

  test("getToken liefert null wenn nichts gespeichert", () => {
    expect(getToken()).toBeNull();
  });

  test("setToken speichert, getToken liest zurueck", async () => {
    await setToken("abc123");
    expect(getToken()).toBe("abc123");
  });
});
```

- [ ] **Step 2: Test ausführen (muss scheitern)**

Run: `npx jest tests/settings.test.ts`
Expected: FAIL ("Cannot find module '../src/lib/settings'").

- [ ] **Step 3: Implementierung schreiben**

```ts
// src/lib/settings.ts
const KEY = "todoistToken";

export function getToken(): string | null {
  const v = Office.context.roamingSettings.get(KEY) as string | undefined;
  return v && v.length > 0 ? v : null;
}

export function setToken(token: string): Promise<void> {
  Office.context.roamingSettings.set(KEY, token);
  return new Promise<void>((resolve, reject) => {
    Office.context.roamingSettings.saveAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}
```

- [ ] **Step 4: Test ausführen (muss passen)**

Run: `npx jest tests/settings.test.ts`
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
git add src/lib/settings.ts tests/settings.test.ts
git commit -m "Settings: Todoist-Token in roamingSettings persistieren"
```

---

### Task 3: EML-Builder (Mail zu MIME)

**Files:**
- Create: `src/lib/emlBuilder.ts`
- Test: `tests/emlBuilder.test.ts`

**Interfaces:**
- Consumes: nichts (reine Logik).
- Produces:
  - `interface MailAttachment { name: string; contentType: string; base64: string; }`
  - `interface MailData { subject: string; from: string; to: string[]; cc: string[]; date: string; htmlBody: string; attachments: MailAttachment[]; }`
  - `buildEml(mail: MailData, boundary?: string): string`
  - `encodeHeader(value: string): string`
  - `utf8ToBase64(str: string): string`

- [ ] **Step 1: Failing tests schreiben**

```ts
// tests/emlBuilder.test.ts
import { buildEml, encodeHeader, utf8ToBase64, MailData } from "../src/lib/emlBuilder";

const base: MailData = {
  subject: "Hallo",
  from: "A B <a@b.ch>",
  to: ["C D <c@d.ch>"],
  cc: [],
  date: "Mon, 01 Jan 2026 10:00:00 +0000",
  htmlBody: "<p>Text</p>",
  attachments: [],
};

describe("encodeHeader", () => {
  test("ASCII bleibt unveraendert", () => {
    expect(encodeHeader("Hello")).toBe("Hello");
  });
  test("Umlaute werden RFC2047-kodiert", () => {
    const out = encodeHeader("Grüezi");
    expect(out.startsWith("=?UTF-8?B?")).toBe(true);
    expect(out.endsWith("?=")).toBe(true);
  });
});

describe("buildEml ohne Anhang", () => {
  const eml = buildEml(base);
  test("enthaelt Header", () => {
    expect(eml).toContain("Subject: Hallo");
    expect(eml).toContain("From: A B <a@b.ch>");
    expect(eml).toContain("To: C D <c@d.ch>");
    expect(eml).toContain("MIME-Version: 1.0");
  });
  test("Body ist base64-kodiert", () => {
    expect(eml).toContain("Content-Transfer-Encoding: base64");
    expect(eml).toContain(utf8ToBase64("<p>Text</p>"));
  });
});

describe("buildEml mit Anhang", () => {
  const eml = buildEml(
    { ...base, attachments: [{ name: "doc.pdf", contentType: "application/pdf", base64: "QUJD" }] },
    "BOUND",
  );
  test("ist multipart/mixed mit Boundary", () => {
    expect(eml).toContain('Content-Type: multipart/mixed; boundary="BOUND"');
    expect(eml).toContain("--BOUND");
    expect(eml).toContain("--BOUND--");
  });
  test("enthaelt Anhang-Teil", () => {
    expect(eml).toContain('Content-Disposition: attachment; filename="doc.pdf"');
    expect(eml).toContain("QUJD");
  });
});
```

- [ ] **Step 2: Test ausführen (muss scheitern)**

Run: `npx jest tests/emlBuilder.test.ts`
Expected: FAIL ("Cannot find module '../src/lib/emlBuilder'").

- [ ] **Step 3: Implementierung schreiben**

```ts
// src/lib/emlBuilder.ts
export interface MailAttachment { name: string; contentType: string; base64: string; }
export interface MailData {
  subject: string;
  from: string;
  to: string[];
  cc: string[];
  date: string;
  htmlBody: string;
  attachments: MailAttachment[];
}

const CRLF = "\r\n";

export function utf8ToBase64(str: string): string {
  return btoa(unescape(encodeURIComponent(str)));
}

export function encodeHeader(value: string): string {
  if (/^[\x00-\x7F]*$/.test(value)) return value;
  return `=?UTF-8?B?${utf8ToBase64(value)}?=`;
}

function chunk76(b64: string): string {
  return b64.replace(/(.{76})/g, `$1${CRLF}`);
}

function htmlPart(html: string): string {
  return (
    `Content-Type: text/html; charset=UTF-8${CRLF}` +
    `Content-Transfer-Encoding: base64${CRLF}${CRLF}` +
    chunk76(utf8ToBase64(html))
  );
}

export function buildEml(mail: MailData, boundary = "=_todoist_eml_boundary_"): string {
  const headers: string[] = [
    `Subject: ${encodeHeader(mail.subject)}`,
    `From: ${mail.from}`,
  ];
  if (mail.to.length) headers.push(`To: ${mail.to.join(", ")}`);
  if (mail.cc.length) headers.push(`Cc: ${mail.cc.join(", ")}`);
  headers.push(`Date: ${mail.date}`);
  headers.push(`MIME-Version: 1.0`);

  if (mail.attachments.length === 0) {
    return headers.join(CRLF) + CRLF + htmlPart(mail.htmlBody) + CRLF;
  }

  headers.push(`Content-Type: multipart/mixed; boundary="${boundary}"`);
  const parts: string[] = [`--${boundary}${CRLF}${htmlPart(mail.htmlBody)}`];
  for (const att of mail.attachments) {
    parts.push(
      `--${boundary}${CRLF}` +
        `Content-Type: ${att.contentType}; name="${att.name}"${CRLF}` +
        `Content-Transfer-Encoding: base64${CRLF}` +
        `Content-Disposition: attachment; filename="${att.name}"${CRLF}${CRLF}` +
        chunk76(att.base64),
    );
  }
  return headers.join(CRLF) + CRLF + CRLF + parts.join(CRLF) + CRLF + `--${boundary}--${CRLF}`;
}
```

- [ ] **Step 4: Test ausführen (muss passen)**

Run: `npx jest tests/emlBuilder.test.ts`
Expected: PASS (6 tests).

- [ ] **Step 5: Commit**

```bash
git add src/lib/emlBuilder.ts tests/emlBuilder.test.ts
git commit -m "EML-Builder: MailData zu gueltigem MIME (.eml)"
```

---

### Task 4: Mail-Reader (Office.js zu MailData)

**Files:**
- Create: `src/lib/mailReader.ts`
- Test: `tests/mailReader.test.ts`

**Interfaces:**
- Consumes: `Office.context.mailbox.item`, `MailData`/`MailAttachment` aus Task 3.
- Produces:
  - `readCurrentMail(): Promise<MailData>`

- [ ] **Step 1: Failing test schreiben**

```ts
// tests/mailReader.test.ts
import { readCurrentMail } from "../src/lib/mailReader";

function mockItem() {
  const O = (global as any).Office;
  (global as any).Office.context.mailbox = {
    item: {
      subject: "Betreff",
      from: { displayName: "Sender", emailAddress: "s@x.ch" },
      to: [{ displayName: "Empf", emailAddress: "e@x.ch" }],
      cc: [],
      dateTimeCreated: new Date("2026-01-01T10:00:00Z"),
      body: {
        getAsync: (_t: string, cb: (r: any) => void) =>
          cb({ status: O.AsyncResultStatus.Succeeded, value: "<p>Hi</p>" }),
      },
      attachments: [
        { id: "1", name: "a.pdf", contentType: "application/pdf", attachmentType: "file" },
        { id: "2", name: "inline.png", contentType: "image/png", attachmentType: "item" },
      ],
      getAttachmentContentAsync: (id: string, cb: (r: any) => void) =>
        cb({
          status: O.AsyncResultStatus.Succeeded,
          value: { format: O.MailboxEnums.AttachmentContentFormat.Base64, content: "B64-" + id },
        }),
    },
  };
}

test("liest Betreff, Adressen, Body und nur File-Anhaenge", async () => {
  mockItem();
  const mail = await readCurrentMail();
  expect(mail.subject).toBe("Betreff");
  expect(mail.from).toBe("Sender <s@x.ch>");
  expect(mail.to).toEqual(["Empf <e@x.ch>"]);
  expect(mail.htmlBody).toBe("<p>Hi</p>");
  expect(mail.attachments).toHaveLength(1);
  expect(mail.attachments[0]).toEqual({ name: "a.pdf", contentType: "application/pdf", base64: "B64-1" });
});
```

- [ ] **Step 2: Test ausführen (muss scheitern)**

Run: `npx jest tests/mailReader.test.ts`
Expected: FAIL ("Cannot find module '../src/lib/mailReader'").

- [ ] **Step 3: Implementierung schreiben**

```ts
// src/lib/mailReader.ts
import { MailData, MailAttachment } from "./emlBuilder";

function fmtAddress(a: { displayName?: string; emailAddress: string }): string {
  return a.displayName ? `${a.displayName} <${a.emailAddress}>` : a.emailAddress;
}

function getHtmlBody(item: any): Promise<string> {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, (r: any) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error),
    );
  });
}

function getAttachmentBase64(item: any, id: string): Promise<string | null> {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(id, (r: any) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return reject(r.error);
      if (r.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) resolve(r.value.content);
      else resolve(null);
    });
  });
}

export async function readCurrentMail(): Promise<MailData> {
  const item: any = Office.context.mailbox.item;
  const htmlBody = await getHtmlBody(item);

  const attachments: MailAttachment[] = [];
  const fileAtts = (item.attachments || []).filter(
    (a: any) => a.attachmentType === Office.MailboxEnums.AttachmentType.File,
  );
  for (const a of fileAtts) {
    try {
      const base64 = await getAttachmentBase64(item, a.id);
      if (base64) attachments.push({ name: a.name, contentType: a.contentType, base64 });
    } catch (e) {
      console.error("Anhang nicht lesbar:", a.name, e);
    }
  }

  return {
    subject: item.subject || "(kein Betreff)",
    from: item.from ? fmtAddress(item.from) : "",
    to: (item.to || []).map(fmtAddress),
    cc: (item.cc || []).map(fmtAddress),
    date: (item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date()).toUTCString(),
    htmlBody,
    attachments,
  };
}
```

- [ ] **Step 4: Test ausführen (muss passen)**

Run: `npx jest tests/mailReader.test.ts`
Expected: PASS (1 test).

- [ ] **Step 5: Commit**

```bash
git add src/lib/mailReader.ts tests/mailReader.test.ts
git commit -m "Mail-Reader: Office.js-Item zu MailData (nur File-Anhaenge)"
```

---

### Task 5: Todoist-API-Client

**Files:**
- Create: `src/lib/todoist.ts`
- Test: `tests/todoist.test.ts`

**Interfaces:**
- Consumes: `fetch`, `FormData`, `Blob`.
- Produces:
  - `interface TodoistTask { id: string; content: string; project_id: string; }`
  - `interface UploadedFile { file_url: string; file_name: string; file_type: string; }`
  - `getTasks(token: string, filter?: string): Promise<TodoistTask[]>`
  - `searchTasks(token: string, query: string): Promise<TodoistTask[]>`
  - `uploadFile(token: string, file: Blob, fileName: string): Promise<UploadedFile>`
  - `addComment(token: string, taskId: string, file: UploadedFile, content?: string): Promise<void>`
  - `class TodoistError extends Error { status: number; body: string; }`

- [ ] **Step 1: Failing tests schreiben**

```ts
// tests/todoist.test.ts
import { getTasks, uploadFile, addComment, TodoistError } from "../src/lib/todoist";

function mockFetch(status: number, json: unknown) {
  (global as any).fetch = jest.fn().mockResolvedValue({
    ok: status >= 200 && status < 300,
    status,
    json: async () => json,
    text: async () => JSON.stringify(json),
  });
}

describe("getTasks", () => {
  test("ruft REST /tasks mit Bearer-Token und Filter", async () => {
    mockFetch(200, [{ id: "1", content: "Test", project_id: "p1" }]);
    const tasks = await getTasks("tok", "(today | overdue)");
    expect(tasks).toHaveLength(1);
    const [url, opts] = (global as any).fetch.mock.calls[0];
    expect(url).toContain("https://api.todoist.com/rest/v2/tasks?filter=");
    expect(opts.headers.Authorization).toBe("Bearer tok");
  });
  test("wirft TodoistError bei 401", async () => {
    mockFetch(401, { error: "unauthorized" });
    await expect(getTasks("bad")).rejects.toBeInstanceOf(TodoistError);
  });
});

describe("uploadFile", () => {
  test("POSTet multipart an sync/uploads/add", async () => {
    mockFetch(200, { file_url: "u", file_name: "m.eml", file_type: "message/rfc822" });
    const out = await uploadFile("tok", new Blob(["x"]), "m.eml");
    expect(out.file_url).toBe("u");
    const [url, opts] = (global as any).fetch.mock.calls[0];
    expect(url).toBe("https://api.todoist.com/sync/v9/uploads/add");
    expect(opts.method).toBe("POST");
    expect(opts.body).toBeInstanceOf(FormData);
  });
});

describe("addComment", () => {
  test("POSTet Kommentar mit Anhang-Objekt", async () => {
    mockFetch(204, {});
    await addComment("tok", "task1", { file_url: "u", file_name: "m.eml", file_type: "message/rfc822" });
    const [url, opts] = (global as any).fetch.mock.calls[0];
    expect(url).toBe("https://api.todoist.com/rest/v2/comments");
    const body = JSON.parse(opts.body);
    expect(body.task_id).toBe("task1");
    expect(body.attachment.file_url).toBe("u");
    expect(body.attachment.resource_type).toBe("file");
  });
});
```

- [ ] **Step 2: Test ausführen (muss scheitern)**

Run: `npx jest tests/todoist.test.ts`
Expected: FAIL ("Cannot find module '../src/lib/todoist'").

- [ ] **Step 3: Implementierung schreiben**

```ts
// src/lib/todoist.ts
const REST = "https://api.todoist.com/rest/v2";
const SYNC = "https://api.todoist.com/sync/v9";

export interface TodoistTask { id: string; content: string; project_id: string; }
export interface UploadedFile { file_url: string; file_name: string; file_type: string; }

export class TodoistError extends Error {
  constructor(public status: number, public body: string) {
    super(`Todoist API ${status}: ${body}`);
    this.name = "TodoistError";
  }
}

function auth(token: string): Record<string, string> {
  return { Authorization: `Bearer ${token}` };
}

async function ensureOk(res: Response): Promise<Response> {
  if (!res.ok) throw new TodoistError(res.status, await res.text());
  return res;
}

export async function getTasks(token: string, filter = "(today | overdue)"): Promise<TodoistTask[]> {
  const res = await fetch(`${REST}/tasks?filter=${encodeURIComponent(filter)}`, { headers: auth(token) });
  return (await ensureOk(res)).json();
}

export async function searchTasks(token: string, query: string): Promise<TodoistTask[]> {
  const res = await fetch(`${REST}/tasks?filter=${encodeURIComponent("search: " + query)}`, { headers: auth(token) });
  return (await ensureOk(res)).json();
}

export async function uploadFile(token: string, file: Blob, fileName: string): Promise<UploadedFile> {
  const form = new FormData();
  form.append("file_name", fileName);
  form.append("file", file, fileName);
  const res = await fetch(`${SYNC}/uploads/add`, { method: "POST", headers: auth(token), body: form });
  return (await ensureOk(res)).json();
}

export async function addComment(token: string, taskId: string, file: UploadedFile, content = ""): Promise<void> {
  const res = await fetch(`${REST}/comments`, {
    method: "POST",
    headers: { ...auth(token), "Content-Type": "application/json" },
    body: JSON.stringify({
      task_id: taskId,
      content: content || file.file_name,
      attachment: { resource_type: "file", file_url: file.file_url, file_name: file.file_name, file_type: file.file_type },
    }),
  });
  await ensureOk(res);
}
```

- [ ] **Step 4: Test ausführen (muss passen)**

Run: `npx jest tests/todoist.test.ts`
Expected: PASS (4 tests).

- [ ] **Step 5: Commit**

```bash
git add src/lib/todoist.ts tests/todoist.test.ts
git commit -m "Todoist-Client: getTasks/searchTasks/uploadFile/addComment"
```

---

### Task 6: Attach-Orchestrierung (Datei-Bau + Upload + Kommentar)

**Files:**
- Create: `src/lib/attachToTask.ts`
- Test: `tests/attachToTask.test.ts`

**Interfaces:**
- Consumes: `readCurrentMail` (T4), `buildEml` (T3), `uploadFile`/`addComment` (T5).
- Produces:
  - `const MAX_BYTES = 26214400` (25 MB)
  - `emlBlobFor(mail: MailData): { blob: Blob; fileName: string }`
  - `attachCurrentMailToTask(token: string, taskId: string): Promise<void>` (wirft `Error` mit Klartext bei >25 MB)

- [ ] **Step 1: Failing tests schreiben**

```ts
// tests/attachToTask.test.ts
import { emlBlobFor, MAX_BYTES } from "../src/lib/attachToTask";
import { MailData } from "../src/lib/emlBuilder";

const mail: MailData = {
  subject: "Re: Angebot / Offerte",
  from: "a@b.ch", to: [], cc: [], date: "Mon, 01 Jan 2026 10:00:00 +0000",
  htmlBody: "<p>x</p>", attachments: [],
};

test("MAX_BYTES ist 25 MB", () => {
  expect(MAX_BYTES).toBe(25 * 1024 * 1024);
});

test("emlBlobFor erzeugt .eml-Dateinamen aus bereinigtem Betreff", () => {
  const { blob, fileName } = emlBlobFor(mail);
  expect(fileName).toBe("Re_ Angebot _ Offerte.eml");
  expect(blob.type).toBe("message/rfc822");
});
```

- [ ] **Step 2: Test ausführen (muss scheitern)**

Run: `npx jest tests/attachToTask.test.ts`
Expected: FAIL ("Cannot find module").

- [ ] **Step 3: Implementierung schreiben**

```ts
// src/lib/attachToTask.ts
import { buildEml, MailData } from "./emlBuilder";
import { readCurrentMail } from "./mailReader";
import { uploadFile, addComment } from "./todoist";

export const MAX_BYTES = 25 * 1024 * 1024;

function sanitizeFileName(subject: string): string {
  const cleaned = (subject || "Mail").replace(/[\\/:*?"<>|]/g, "_").trim();
  return `${cleaned.slice(0, 120) || "Mail"}.eml`;
}

export function emlBlobFor(mail: MailData): { blob: Blob; fileName: string } {
  const eml = buildEml(mail);
  const blob = new Blob([eml], { type: "message/rfc822" });
  return { blob, fileName: sanitizeFileName(mail.subject) };
}

export async function attachCurrentMailToTask(token: string, taskId: string): Promise<void> {
  const mail = await readCurrentMail();
  const { blob, fileName } = emlBlobFor(mail);
  if (blob.size > MAX_BYTES) {
    throw new Error(`Mail ist ${(blob.size / 1024 / 1024).toFixed(1)} MB gross, Todoist erlaubt max 25 MB.`);
  }
  const uploaded = await uploadFile(token, blob, fileName);
  await addComment(token, taskId, uploaded);
}
```

- [ ] **Step 4: Test ausführen (muss passen)**

Run: `npx jest tests/attachToTask.test.ts`
Expected: PASS (2 tests).

- [ ] **Step 5: Commit**

```bash
git add src/lib/attachToTask.ts tests/attachToTask.test.ts
git commit -m "Attach-Orchestrierung: .eml bauen, Groesse pruefen, hochladen, kommentieren"
```

---

### Task 7: Task-Pane-UI verdrahten

**Files:**
- Modify: `src/taskpane/taskpane.html` (UI-Struktur)
- Modify: `src/taskpane/taskpane.ts` (Logik)
- Modify: `src/taskpane/taskpane.css` (minimal)

**Interfaces:**
- Consumes: `getToken`/`setToken` (T2), `getTasks`/`searchTasks` (T5), `attachCurrentMailToTask` (T6).
- Produces: lauffähiges Pane. Keine neuen Exporte.

- [ ] **Step 1: HTML-Grundgerüst setzen**

`src/taskpane/taskpane.html` Body-Inhalt ersetzen durch:
```html
<main id="app" style="padding:12px; font-family:Segoe UI, sans-serif;">
  <section id="token-section" hidden>
    <p>Todoist-API-Token (Todoist: Einstellungen &gt; Integrationen &gt; Entwickler):</p>
    <input id="token-input" type="password" style="width:100%" />
    <button id="token-save">Speichern</button>
  </section>
  <section id="task-section" hidden>
    <input id="search" placeholder="Task suchen..." style="width:100%; margin-bottom:8px" />
    <ul id="task-list" style="list-style:none; padding:0; margin:0"></ul>
  </section>
  <p id="status" role="status" style="margin-top:8px"></p>
</main>
```

- [ ] **Step 2: Pane-Logik schreiben**

`src/taskpane/taskpane.ts` ersetzen durch:
```ts
import { getToken, setToken } from "../lib/settings";
import { getTasks, searchTasks, TodoistTask } from "../lib/todoist";
import { attachCurrentMailToTask } from "../lib/attachToTask";

let busy = false;

function $(id: string): HTMLElement { return document.getElementById(id)!; }
function setStatus(msg: string, isError = false): void {
  const el = $("status");
  el.textContent = msg;
  el.style.color = isError ? "#b00" : "#060";
  if (isError) console.error(msg);
}

function renderTasks(tasks: TodoistTask[], token: string): void {
  const list = $("task-list");
  list.innerHTML = "";
  if (tasks.length === 0) { setStatus("Keine Tasks gefunden."); return; }
  for (const t of tasks) {
    const li = document.createElement("li");
    const btn = document.createElement("button");
    btn.textContent = t.content;
    btn.style.cssText = "width:100%; text-align:left; padding:8px; margin:2px 0; cursor:pointer";
    btn.onclick = () => attach(token, t);
    li.appendChild(btn);
    list.appendChild(li);
  }
}

async function attach(token: string, task: TodoistTask): Promise<void> {
  if (busy) return;
  busy = true;
  setStatus(`Haenge Mail an "${task.content}" ...`);
  try {
    await attachCurrentMailToTask(token, task.id);
    setStatus(`Erledigt: Mail haengt an "${task.content}".`);
  } catch (e) {
    setStatus(`Fehler: ${(e as Error).message}`, true);
  } finally {
    busy = false;
  }
}

async function loadTasks(token: string): Promise<void> {
  try {
    const tasks = await getTasks(token);
    renderTasks(tasks, token);
  } catch (e) {
    setStatus(`Token ungueltig oder Abruf fehlgeschlagen: ${(e as Error).message}`, true);
    showTokenSection();
  }
}

function showTokenSection(): void {
  $("token-section").hidden = false;
  $("task-section").hidden = true;
}

function showTaskSection(): void {
  $("token-section").hidden = true;
  $("task-section").hidden = false;
}

function wireSearch(token: string): void {
  let timer: number | undefined;
  ($("search") as HTMLInputElement).addEventListener("input", (e) => {
    const q = (e.target as HTMLInputElement).value.trim();
    window.clearTimeout(timer);
    timer = window.setTimeout(async () => {
      try {
        const tasks = q ? await searchTasks(token, q) : await getTasks(token);
        renderTasks(tasks, token);
      } catch (err) {
        setStatus(`Suche fehlgeschlagen: ${(err as Error).message}`, true);
      }
    }, 300);
  });
}

Office.onReady(() => {
  ($("token-save") as HTMLButtonElement).onclick = async () => {
    const val = ($("token-input") as HTMLInputElement).value.trim();
    if (!val) return;
    await setToken(val);
    start();
  };
  start();
});

function start(): void {
  const token = getToken();
  if (!token) { showTokenSection(); return; }
  showTaskSection();
  wireSearch(token);
  loadTasks(token);
}
```

- [ ] **Step 3: Build prüfen**

Run: `npm run build`
Expected: Webpack-Build ohne Fehler (TypeScript kompiliert).

- [ ] **Step 4: Unit-Suite gesamthaft prüfen**

Run: `npx jest`
Expected: alle Tests PASS (Tasks 1-6).

- [ ] **Step 5: Commit**

```bash
git add src/taskpane/
git commit -m "Task-Pane-UI: Token-Flow, Task-Liste, Suche, Anhaengen mit Status"
```

---

### Task 8: Manifest + lokaler Sideload-Test (Dev-Tenant)

**Files:**
- Modify: `manifest.xml`
- Create: `docs/SIDELOAD.md`

**Interfaces:**
- Consumes: gehostete/lokale Pane-URL.
- Produces: gültiges Manifest mit Mailbox 1.8 + ReadItem + Button.

- [ ] **Step 1: Manifest auf Mailbox 1.8 + Button setzen**

In `manifest.xml` sicherstellen:
- `<Requirements><Sets><Set Name="Mailbox" MinVersion="1.8"/></Sets></Requirements>`
- `<Permissions>ReadItem</Permissions>`
- Ein `MessageReadCommandSurface`-Button, der die Task-Pane öffnet (Standard vom Scaffold, Beschriftung „An Todoist-Task").
- Eindeutige `<Id>` (GUID), `<DisplayName DefaultValue="Mail an Todoist"/>`.

- [ ] **Step 2: Dev-Zertifikat + lokalen Server starten**

```bash
npx office-addin-dev-certs install
npm run start
```
Erwartet: Dev-Server auf `https://localhost:3000`, `office-addin-debugging` versucht zu sideloaden.

- [ ] **Step 3: Manuell im Microsoft-365-Developer-Tenant sideloaden**

In `docs/SIDELOAD.md` dokumentieren (und durchführen): Outlook im Web des Dev-Tenants &gt; Apps &gt; Apps verwalten &gt; Benutzerdefiniertes Add-In &gt; „Aus Datei" &gt; `manifest.xml`. (Der CMI-Produktiv-Tenant erlaubt das nicht, daher Dev-Tenant.)

- [ ] **Step 4: End-to-End manuell verifizieren**

Mail öffnen &gt; Button &gt; Token eingeben &gt; Task anklicken &gt; in Todoist prüfen, dass die `.eml` als Kommentar-Anhang am Task hängt und sich öffnen lässt.
Erwartet: Anhang vorhanden, öffnet in Outlook, Betreff/Body/Anhänge enthalten.

- [ ] **Step 5: Commit**

```bash
git add manifest.xml docs/SIDELOAD.md
git commit -m "Manifest: Mailbox 1.8 + ReadItem + Button; Sideload-Anleitung"
```

---

### Task 9: Produktion (GitHub Pages) + IT-Handoff

**Files:**
- Create: `.github/workflows/deploy.yml` (oder manueller Pages-Build)
- Create: `manifest.prod.xml`
- Create: `docs/IT-ROLLOUT.md`

**Interfaces:**
- Consumes: gebaute statische Dateien.
- Produces: öffentliche HTTPS-URL + Produktiv-Manifest + IT-Anleitung.

- [ ] **Step 1: Produktions-Build erzeugen**

Run: `npm run build`
Expected: `dist/` mit statischen Assets.

- [ ] **Step 2: GitHub Pages aufsetzen**

Repo auf GitHub anlegen (`manuelweingartner/outlook-todoist-addin`), Pages auf `gh-pages` oder `/docs`-Build. `.github/workflows/deploy.yml` baut und published `dist/`.
```yaml
name: deploy
on: { push: { branches: [master] } }
permissions: { contents: write }
jobs:
  build:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4
      - uses: actions/setup-node@v4
        with: { node-version: "20" }
      - run: npm ci && npm run build
      - uses: peaceiris/actions-gh-pages@v4
        with:
          github_token: ${{ secrets.GITHUB_TOKEN }}
          publish_dir: ./dist
```

- [ ] **Step 3: `manifest.prod.xml` mit Pages-URLs**

Kopie von `manifest.xml`, alle `https://localhost:3000`-URLs ersetzt durch `https://manuelweingartner.github.io/outlook-todoist-addin/`.

- [ ] **Step 4: IT-Handoff dokumentieren**

`docs/IT-ROLLOUT.md`:
```markdown
# IT-Rollout: Add-in „Mail an Todoist"
Bitte zentral bereitstellen (M365 Admin-Center > Einstellungen > Integrierte Apps > Benutzerdefinierte Apps hochladen), genau wie bei "CMI Mail".
- Manifest-URL: https://manuelweingartner.github.io/outlook-todoist-addin/manifest.prod.xml
- Benoetigte Berechtigung: NUR Office.js "ReadItem" (im Manifest). KEIN Microsoft Graph, KEIN Mail.Read, KEIN Azure-App, KEIN Mailbox-Lese-Consent.
- Datenfluss: Add-in liest nur die aktuell geoeffnete Mail lokal im Client und laedt sie zum persoenlichen Todoist-Konto des Nutzers hoch. Kein Server von Drittanbietern ausser Todoist.
- Zuweisung: nur an Manuel Weingartner (bzw. definierte Gruppe).
```

- [ ] **Step 5: Commit + Push**

```bash
git add .github/ manifest.prod.xml docs/IT-ROLLOUT.md
git commit -m "Produktion: GitHub-Pages-Deploy, Prod-Manifest, IT-Rollout-Doku"
git push -u origin master
```

---

## Self-Review

**Spec coverage:**
- Ziel-Ablauf (Button, Liste, Klick, Anhängen) → Tasks 6, 7. ✓
- Mail-Extraktion ohne Graph (Office.js, .eml-Rekonstruktion) → Tasks 3, 4. ✓
- Todoist-Token-Auth + roamingSettings → Task 2. ✓
- Upload + Kommentar, CORS, kein Proxy → Task 5. ✓
- Task-Auswahl (Heute/Überfällig + Suche) → Tasks 5, 7. ✓
- Fehlerbehandlung (401, Anhang-Fehler, 25 MB, Doppelklick, kein stilles catch) → Tasks 4, 5, 6, 7. ✓
- Tests (EML-Builder, Payloads, Persistenz, gemockt) → Tasks 2-6. ✓
- Hosting GitHub Pages + Manifest XML + IT-Anfrage (nur Deploy) → Tasks 8, 9. ✓
- Manueller E2E-Test im Dev-Tenant → Task 8. ✓

**Placeholder scan:** Keine TBD/TODO; alle Code-Schritte mit echtem Code; Manifest-Bearbeitung in Task 8 als konkrete Feldliste (Scaffold liefert XML-Grundgerüst).

**Type consistency:** `MailData`/`MailAttachment` (T3) konsistent in T4/T6 verwendet. `TodoistTask`/`UploadedFile` (T5) konsistent in T6/T7. `attachCurrentMailToTask(token, taskId)` Signatur in T6 definiert, in T7 so aufgerufen. `getToken`/`setToken` (T2) in T7 so genutzt.
