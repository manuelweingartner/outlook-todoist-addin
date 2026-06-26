const REST = "https://api.todoist.com/rest/v2";
const SYNC = "https://api.todoist.com/sync/v9";

export interface TodoistTask { id: string; content: string; project_id: string; }
export interface UploadedFile { file_url: string; file_name: string; file_type: string; }

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
