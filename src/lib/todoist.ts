// Todoist Unified API v1. Die alten Endpoints (REST v2 / Sync v9) wurden von
// Todoist abgeschaltet (HTTP 410 Gone, ohne CORS-Header -> Browser meldet
// "Failed to fetch"). Alles laeuft jetzt ueber api.todoist.com/api/v1.
const API = "https://api.todoist.com/api/v1";

export interface TodoistTask {
  id: string;
  content: string;
  project_id: string;
  priority?: number;                                  // 1 bis 4, 4 = hoechste (P1)
  due?: { date: string; datetime?: string } | null;
}
export interface TodoistProject { id: string; name: string; }
export interface UploadedFile { file_url: string; file_name: string; file_type: string; }

// v1-Listen sind paginiert: { results: [...], next_cursor: ... }.
interface Paginated<T> { results: T[]; next_cursor?: string | null; }

export class TodoistError extends Error {
  public status: number;
  public body: string;

  constructor(status: number, body: string) {
    super(`Todoist API ${status}: ${body}`);
    this.name = "TodoistError";
    this.status = status;
    this.body = body;
    Object.setPrototypeOf(this, TodoistError.prototype);
  }
}

function auth(token: string): Record<string, string> {
  return { Authorization: `Bearer ${token}` };
}

async function ensureOk(res: Response): Promise<Response> {
  if (!res.ok) throw new TodoistError(res.status, await res.text());
  return res;
}

// Akzeptiert sowohl die paginierte v1-Form als auch (defensiv) ein nacktes Array.
function unwrap<T>(data: Paginated<T> | T[]): T[] {
  return Array.isArray(data) ? data : data.results ?? [];
}

async function tasksByQuery(token: string, query: string): Promise<TodoistTask[]> {
  const res = await fetch(`${API}/tasks/filter?query=${encodeURIComponent(query)}`, { headers: auth(token) });
  return unwrap<TodoistTask>(await (await ensureOk(res)).json());
}

export async function getTasks(token: string, query = "(today | overdue)"): Promise<TodoistTask[]> {
  return tasksByQuery(token, query);
}

export async function searchTasks(token: string, query: string): Promise<TodoistTask[]> {
  return tasksByQuery(token, `search: ${query}`);
}

export async function uploadFile(token: string, file: Blob, fileName: string): Promise<UploadedFile> {
  const form = new FormData();
  form.append("file_name", fileName);
  form.append("file", file, fileName);
  const res = await fetch(`${API}/uploads`, { method: "POST", headers: auth(token), body: form });
  return (await ensureOk(res)).json();
}

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
