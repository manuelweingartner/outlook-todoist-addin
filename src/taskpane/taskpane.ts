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
  ($("new-task") as HTMLButtonElement).disabled = tooLarge();
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
