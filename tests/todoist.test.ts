import { getTasks, uploadFile, addComment, searchTasks, TodoistError } from "../src/lib/todoist";

function mockFetch(status: number, json: unknown) {
  (global as any).fetch = jest.fn().mockResolvedValue({
    ok: status >= 200 && status < 300,
    status,
    json: async () => json,
    text: async () => JSON.stringify(json),
  });
}

describe("getTasks", () => {
  test("ruft API v1 /tasks/filter mit Bearer-Token und query", async () => {
    mockFetch(200, { results: [{ id: "1", content: "Test", project_id: "p1" }], next_cursor: null });
    const tasks = await getTasks("tok", "(today | overdue)");
    expect(tasks).toHaveLength(1);
    expect(tasks[0].content).toBe("Test");
    const [url, opts] = (global as any).fetch.mock.calls[0];
    expect(url).toContain("https://api.todoist.com/api/v1/tasks/filter?query=");
    expect(opts.headers.Authorization).toBe("Bearer tok");
  });
  test("kommt auch mit nacktem Array klar (defensiv)", async () => {
    mockFetch(200, [{ id: "1", content: "Test", project_id: "p1" }]);
    const tasks = await getTasks("tok");
    expect(tasks).toHaveLength(1);
  });
  test("wirft TodoistError bei 401", async () => {
    mockFetch(401, { error: "unauthorized" });
    await expect(getTasks("bad")).rejects.toBeInstanceOf(TodoistError);
  });
});

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

describe("searchTasks", () => {
  test("nutzt search:-query gegen api/v1/tasks/filter und Bearer-Token", async () => {
    mockFetch(200, { results: [{ id: "2", content: "Mail X", project_id: "p1" }], next_cursor: null });
    await searchTasks("tok", "Mail");
    const [url, opts] = (global as any).fetch.mock.calls[0];
    expect(url).toContain("https://api.todoist.com/api/v1/tasks/filter?query=");
    expect(decodeURIComponent(url)).toContain("search: Mail");
    expect(opts.headers.Authorization).toBe("Bearer tok");
  });
});
