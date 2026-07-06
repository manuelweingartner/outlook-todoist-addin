import { uploadFile, addComment, TodoistError, deleteComment, getProjects, createTask, getAllTasks, isAuthError, getOldTaskIds } from "../src/lib/todoist";

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
  test("POSTet Options als JSON-Body", async () => {
    mockFetch(200, { id: "9", content: "Neu", project_id: "p1" });
    const task = await createTask("tok", { content: "Neu", project_id: "p1", priority: 3, due_date: "2026-07-07" });
    expect(task.id).toBe("9");
    const [url, opts] = (global as any).fetch.mock.calls[0];
    expect(url).toBe("https://api.todoist.com/api/v1/tasks");
    expect(opts.method).toBe("POST");
    expect(JSON.parse(opts.body)).toEqual({ content: "Neu", project_id: "p1", priority: 3, due_date: "2026-07-07" });
  });
  test("minimaler Aufruf nur mit content", async () => {
    mockFetch(200, { id: "9", content: "Neu", project_id: "p1" });
    await createTask("tok", { content: "Neu" });
    expect(JSON.parse((global as any).fetch.mock.calls[0][1].body)).toEqual({ content: "Neu" });
  });
});

describe("isAuthError", () => {
  test("401 und 403 sind Auth-Fehler", () => {
    expect(isAuthError(new TodoistError(401, "unauthorized"))).toBe(true);
    expect(isAuthError(new TodoistError(403, "forbidden"))).toBe(true);
  });
  test("andere Fehler nicht", () => {
    expect(isAuthError(new TodoistError(500, "boom"))).toBe(false);
    expect(isAuthError(new TypeError("Failed to fetch"))).toBe(false);
    expect(isAuthError(undefined)).toBe(false);
  });
});

describe("getAllTasks maxPages", () => {
  test("bricht nach 50 Seiten mit stabilem Cursor ab und liefert Teilmenge", async () => {
    (global as any).fetch = jest.fn().mockResolvedValue({
      ok: true, status: 200,
      json: async () => ({ results: [{ id: "x", content: "T", project_id: "p" }], next_cursor: "immerGleich" }),
      text: async () => "",
    });
    const errSpy = jest.spyOn(console, "error").mockImplementation(() => {});
    const tasks = await getAllTasks("tok");
    expect((global as any).fetch).toHaveBeenCalledTimes(50);
    expect(tasks).toHaveLength(50);
    expect(errSpy).toHaveBeenCalled();
    errSpy.mockRestore();
  });
});

describe("getOldTaskIds", () => {
  test("liest /id_mappings/tasks/<ids> und liefert Map neue-Id -> alte Id", async () => {
    mockFetch(200, [
      { old_id: "918273645", new_id: "6VfWjjjFg2xqX6Pa" },
      { old_id: "111222333", new_id: "6WMVPf8Hm8JP6mC8" },
    ]);
    const map = await getOldTaskIds("tok", ["6VfWjjjFg2xqX6Pa", "6WMVPf8Hm8JP6mC8"]);
    expect(map).toEqual({ "6VfWjjjFg2xqX6Pa": "918273645", "6WMVPf8Hm8JP6mC8": "111222333" });
    const [url, opts] = (global as any).fetch.mock.calls[0];
    expect(url).toBe("https://api.todoist.com/api/v1/id_mappings/tasks/6VfWjjjFg2xqX6Pa,6WMVPf8Hm8JP6mC8");
    expect(opts.headers.Authorization).toBe("Bearer tok");
  });

  test("ueberspringt Eintraege ohne old_id oder new_id", async () => {
    mockFetch(200, [
      { old_id: null, new_id: "a1" },
      { old_id: "9", new_id: null },
      { old_id: "8", new_id: "b2" },
    ]);
    const map = await getOldTaskIds("tok", ["a1", "b2"]);
    expect(map).toEqual({ b2: "8" });
  });

  test("chunked bei mehr als 100 Ids in mehrere Aufrufe", async () => {
    (global as any).fetch = jest.fn().mockResolvedValue({
      ok: true, status: 200, json: async () => [], text: async () => "",
    });
    const ids = Array.from({ length: 250 }, (_, i) => `id${i}`);
    await getOldTaskIds("tok", ids);
    expect((global as any).fetch).toHaveBeenCalledTimes(3);
    const firstUrl = (global as any).fetch.mock.calls[0][0] as string;
    expect(firstUrl.split(",")).toHaveLength(100);
  });

  test("liefert bei leerer Id-Liste {} ohne Fetch", async () => {
    (global as any).fetch = jest.fn();
    const map = await getOldTaskIds("tok", []);
    expect(map).toEqual({});
    expect((global as any).fetch).not.toHaveBeenCalled();
  });

  test("faellt bei API-Fehler auf {} zurueck und loggt (Deep-Link nutzt dann die neue Id)", async () => {
    mockFetch(500, { error: "boom" });
    const errSpy = jest.spyOn(console, "error").mockImplementation(() => {});
    const map = await getOldTaskIds("tok", ["a1"]);
    expect(map).toEqual({});
    expect(errSpy).toHaveBeenCalled();
    errSpy.mockRestore();
  });

  test("faellt bei Netzwerkfehler auf {} zurueck und loggt", async () => {
    (global as any).fetch = jest.fn().mockRejectedValue(new TypeError("Failed to fetch"));
    const errSpy = jest.spyOn(console, "error").mockImplementation(() => {});
    const map = await getOldTaskIds("tok", ["a1"]);
    expect(map).toEqual({});
    expect(errSpy).toHaveBeenCalled();
    errSpy.mockRestore();
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
