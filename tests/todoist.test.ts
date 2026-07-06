import { uploadFile, addComment, TodoistError, deleteComment, getProjects, createTask, getAllTasks } from "../src/lib/todoist";

function mockFetch(status: number, json: unknown) {
  (global as any).fetch = jest.fn().mockResolvedValue({
    ok: status >= 200 && status < 300,
    status,
    json: async () => json,
    text: async () => JSON.stringify(json),
  });
}

describe("uploadFile", () => {
  test("POSTet multipart an api/v1/uploads", async () => {
    mockFetch(200, { file_url: "u", file_name: "m.eml", file_type: "message/rfc822" });
    const out = await uploadFile("tok", new Blob(["x"]), "m.eml");
    expect(out.file_url).toBe("u");
    const [url, opts] = (global as any).fetch.mock.calls[0];
    expect(url).toBe("https://api.todoist.com/api/v1/uploads");
    expect(opts.method).toBe("POST");
    expect(opts.body).toBeInstanceOf(FormData);
    expect(opts.headers.Authorization).toBe("Bearer tok");
  });
});

describe("addComment", () => {
  test("POSTet Kommentar mit Anhang-Objekt an api/v1/comments", async () => {
    mockFetch(204, {});
    await addComment("tok", "task1", { file_url: "u", file_name: "m.eml", file_type: "message/rfc822" });
    const [url, opts] = (global as any).fetch.mock.calls[0];
    expect(url).toBe("https://api.todoist.com/api/v1/comments");
    const body = JSON.parse(opts.body);
    expect(body.task_id).toBe("task1");
    expect(body.attachment.file_url).toBe("u");
    expect(body.attachment.resource_type).toBe("file");
    expect(opts.headers.Authorization).toBe("Bearer tok");
    expect(opts.headers["Content-Type"]).toBe("application/json");
  });
});

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

  test("kommt auch mit nacktem Array klar (defensiv)", async () => {
    (global as any).fetch = jest.fn().mockResolvedValue({
      ok: true, status: 200, json: async () => [{ id: "1", content: "Test", project_id: "p1" }], text: async () => "",
    });
    const tasks = await getAllTasks("tok");
    expect(tasks).toHaveLength(1);
  });
});
