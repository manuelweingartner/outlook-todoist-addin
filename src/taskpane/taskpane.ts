import { getToken, setToken } from "../lib/settings";
import { getAllTasks, getProjects, createTask, deleteComment, isAuthError, TodoistTask } from "../lib/todoist";
import { readAndPrepareCurrentMail, prepareMail, attachPreparedToTask, PreparedMail, MAX_BYTES } from "../lib/attachToTask";
import { MailData } from "../lib/emlBuilder";
import { groupTasks, priorityColor, taskDeepLink, todayIso, filterTasks, dueTodayOrOverdue, extractMailKeywords, suggestTasks, moveSelection, buildNewTaskOptions, DueChip } from "./taskLogic";

let busy = false;
let searchWired = false;
let prepared: PreparedMail | null = null;
let mailData: MailData | null = null;
let projectNames: Record<string, string> = {};
let allTasks: TodoistTask[] = [];
let selectedIndex = 0;

function $(id: string): HTMLElement { return document.getElementById(id)!; }

function openExternal(url: string): void {
  window.open(url, "_blank", "noopener");
}

function visibleRows(): HTMLButtonElement[] {
  return Array.from($("task-groups").querySelectorAll<HTMLButtonElement>(".task-row"));
}

function applySelection(): void {
  const rows = visibleRows();
  rows.forEach((r, i) => r.classList.toggle("selected", i === selectedIndex));
  const sel = rows[selectedIndex];
  if (sel) sel.scrollIntoView({ block: "nearest" });
}

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

function tooLarge(): boolean {
  return !!prepared && prepared.sizeBytes > MAX_BYTES;
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

  const main = document.createElement("span");
  main.className = "task-main";

  const content = document.createElement("span");
  content.className = "content";
  content.textContent = task.content;
  main.appendChild(content);

  const projectName = projectNames[task.project_id];
  if (projectName) {
    const proj = document.createElement("span");
    proj.className = "project";
    proj.textContent = projectName;
    main.appendChild(proj);
  }

  btn.appendChild(main);

  const state = document.createElement("span");
  state.className = "state";
  btn.appendChild(state);

  btn.onclick = () => attach(task, btn, state, li);

  const open = document.createElement("button");
  open.className = "row-open";
  open.title = "In Todoist öffnen";
  open.textContent = "↗"; // Pfeil nach oben rechts
  open.onclick = (e) => { e.stopPropagation(); openExternal(taskDeepLink(task.id)); };

  li.appendChild(btn);
  li.appendChild(open);
  return li;
}

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

  selectedIndex = 0;
  applySelection();
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

// Rendert passend zum aktuellen Suchfeld-Inhalt. Fixt nebenbei: Tippen waehrend
// des Initial-Loads wird beim Load-Ende nicht mehr von der Standardansicht ueberschrieben.
function rerender(): void {
  const q = ($("search") as HTMLInputElement).value.trim();
  if (q) renderSections([["Treffer", filterTasks(allTasks, q, projectNames)]], "Keine Treffer.");
  else renderDefaultView();
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
    renderUndo(li, token, commentId, state, task.id);
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

function renderUndo(li: HTMLElement, token: string, commentId: string, state: HTMLElement, taskId: string): void {
  if (!commentId) return;
  const undo = document.createElement("button");
  undo.className = "undo";
  undo.textContent = "Rückgängig";
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
      setStatus(`Rückgängig fehlgeschlagen: ${(e as Error).message}`, "err", e);
    }
  };
  li.appendChild(undo);
  const openLink = document.createElement("button");
  openLink.className = "undo";
  openLink.textContent = "In Todoist öffnen";
  openLink.onclick = () => openExternal(taskDeepLink(taskId));
  li.appendChild(openLink);
}

async function loadTasks(token: string): Promise<void> {
  $("load-error").hidden = true;
  setSkeleton(true);
  try {
    const [tasks] = await Promise.all([
      getAllTasks(token),
      loadProjects(token),
    ]);
    allTasks = tasks;
    rerender();
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
  ($("search") as HTMLInputElement).addEventListener("input", () => {
    rerender();
  });

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
}

function wireRetry(): void {
  ($("load-error-retry") as HTMLButtonElement).onclick = () => {
    const token = getToken();
    if (token) void loadTasks(token);
  };
}

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
  if (sel.options.length === 0) {
    // Projekte (noch) nicht geladen: leerer value -> projectId null -> Todoist
    // legt in der Inbox an. So ist das Dropdown nie leer, Verhalten bleibt ehrlich.
    const fallback = document.createElement("option");
    fallback.value = "";
    fallback.textContent = "Inbox";
    sel.appendChild(fallback);
  }
}

function resetNewTaskForm(): void {
  ntPriority = 1;
  ntChip = "none";
  ($("nt-due-text") as HTMLInputElement).value = "";
  const prioDefault = $("nt-prio").querySelector<HTMLButtonElement>('button[data-p="1"]');
  if (prioDefault) markActive($("nt-prio"), prioDefault);
  const chipDefault = $("nt-chips").querySelector<HTMLButtonElement>('button[data-chip="none"]');
  if (chipDefault) markActive($("nt-chips"), chipDefault);
}

function wireNewTask(): void {
  const btn = $("new-task") as HTMLButtonElement;
  const form = $("new-task-form") as HTMLFormElement;
  btn.hidden = false;

  btn.onclick = async () => {
    if (tooLarge() || !prepared) return;
    form.hidden = !form.hidden;
    if (!form.hidden) {
      ($("nt-title") as HTMLInputElement).value = prepared.subject || "";
      if (Object.keys(projectNames).length === 0) {
        const token = getToken();
        if (token) await loadProjects(token);
      }
      fillProjectSelect();
      resetNewTaskForm();
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
    ($("nt-cancel") as HTMLButtonElement).disabled = true;
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
      ($("nt-cancel") as HTMLButtonElement).disabled = false;
    }
  };
}

async function start(): Promise<void> {
  const token = getToken();
  if (!token) { showTokenSection(); return; }
  showTaskSection();

  try {
    const { mail, prepared: p } = await readAndPrepareCurrentMail();
    mailData = mail;
    prepared = p;
    renderContextBar();
    renderSizeWarning();
  } catch (e) {
    setStatus(`Mail konnte nicht gelesen werden: ${(e as Error).message}`, "err", e);
  }

  if (!searchWired) { wireSearch(); wireNewTask(); wireRetry(); searchWired = true; }
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
