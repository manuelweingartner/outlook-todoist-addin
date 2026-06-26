import { getToken, setToken } from "../lib/settings";
import { getTasks, searchTasks, TodoistTask } from "../lib/todoist";
import { attachCurrentMailToTask } from "../lib/attachToTask";

let busy = false;
let searchWired = false;

function $(id: string): HTMLElement { return document.getElementById(id)!; }
function setStatus(msg: string, isError = false, cause?: unknown): void {
  const el = $("status");
  el.textContent = msg;
  el.style.color = isError ? "#b00" : "#060";
  if (isError) {
    if (cause !== undefined) console.error(msg, cause);
    else console.error(msg);
  }
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
    setStatus(`Fehler: ${(e as Error).message}`, true, e);
  } finally {
    busy = false;
  }
}

async function loadTasks(token: string): Promise<void> {
  try {
    const tasks = await getTasks(token);
    renderTasks(tasks, token);
  } catch (e) {
    setStatus(`Token ungueltig oder Abruf fehlgeschlagen: ${(e as Error).message}`, true, e);
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

function wireSearch(): void {
  let timer: number | undefined;
  ($("search") as HTMLInputElement).addEventListener("input", (e) => {
    const q = (e.target as HTMLInputElement).value.trim();
    const token = getToken();
    if (!token) return;
    window.clearTimeout(timer);
    timer = window.setTimeout(async () => {
      try {
        const tasks = q ? await searchTasks(token, q) : await getTasks(token);
        renderTasks(tasks, token);
      } catch (err) {
        setStatus(`Suche fehlgeschlagen: ${(err as Error).message}`, true, err);
      }
    }, 300);
  });
}

Office.onReady(() => {
  ($("token-save") as HTMLButtonElement).onclick = async () => {
    const val = ($("token-input") as HTMLInputElement).value.trim();
    if (!val) {
      setStatus("Bitte Token eingeben.", true);
      return;
    }
    await setToken(val);
    start();
  };
  start();
});

function start(): void {
  const token = getToken();
  if (!token) { showTokenSection(); return; }
  showTaskSection();
  if (!searchWired) {
    wireSearch();
    searchWired = true;
  }
  loadTasks(token);
}
