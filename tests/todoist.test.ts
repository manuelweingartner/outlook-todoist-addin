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
