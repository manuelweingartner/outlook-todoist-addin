import { buildSummaryPrompt, parseSummary, summarizeMail, SummaryError } from "../src/lib/summarize";
import { MailData } from "../src/lib/emlBuilder";

const mail: MailData = {
  subject: "Budget Q2",
  from: "Patrick Nanzer <patrick@example.com>",
  to: ["Manuel Weingartner <manuel@example.com>"],
  cc: ["Chef <chef@example.com>"],
  date: "2026-07-07T09:00:00Z",
  htmlBody: "<p>egal</p>",
  attachments: [],
};

function mockFetch(status: number, json: unknown) {
  (global as any).fetch = jest.fn().mockResolvedValue({
    ok: status >= 200 && status < 300,
    status,
    json: async () => json,
    text: async () => JSON.stringify(json),
  });
}

describe("buildSummaryPrompt", () => {
  test("enthaelt Absender, Empfaenger, Betreff und Body-Ausschnitt", () => {
    const p = buildSummaryPrompt(mail, "Bitte gib das Q2-Budget bis Freitag frei.");
    expect(p).toContain("Patrick Nanzer");
    expect(p).toContain("Budget Q2");
    expect(p).toContain("Q2-Budget bis Freitag");
    expect(p).toContain("Manuel Weingartner"); // Empfaenger
  });
  test("verlangt genau einen deutschen Satz ohne Vorspann", () => {
    const p = buildSummaryPrompt(mail, "x").toLowerCase();
    expect(p).toContain("ein");
    expect(p).toContain("satz");
  });
});

describe("parseSummary", () => {
  test("extrahiert den Text aus content[0].text und trimmt", () => {
    expect(parseSummary({ content: [{ type: "text", text: "  Ein Satz.  " }], stop_reason: "end_turn" }))
      .toBe("Ein Satz.");
  });
  test("entfernt umschliessende Anfuehrungszeichen", () => {
    expect(parseSummary({ content: [{ type: "text", text: '"Ein Satz."' }], stop_reason: "end_turn" }))
      .toBe("Ein Satz.");
  });
  test("wirft bei Refusal", () => {
    expect(() => parseSummary({ content: [], stop_reason: "refusal" })).toThrow(SummaryError);
  });
  test("wirft bei leerem/fehlendem Text", () => {
    expect(() => parseSummary({ content: [{ type: "text", text: "   " }], stop_reason: "end_turn" })).toThrow(SummaryError);
    expect(() => parseSummary({ content: [] })).toThrow(SummaryError);
    expect(() => parseSummary({})).toThrow(SummaryError);
  });
});

describe("summarizeMail", () => {
  test("POSTet an die Messages-API mit den richtigen Headern und liefert den Satz", async () => {
    mockFetch(200, { content: [{ type: "text", text: "Patrick bittet um Freigabe des Q2-Budgets bis Freitag." }], stop_reason: "end_turn" });
    const out = await summarizeMail("sk-ant-xyz", mail, "Bitte Q2-Budget freigeben.");
    expect(out).toBe("Patrick bittet um Freigabe des Q2-Budgets bis Freitag.");
    const [url, opts] = (global as any).fetch.mock.calls[0];
    expect(url).toBe("https://api.anthropic.com/v1/messages");
    expect(opts.method).toBe("POST");
    expect(opts.headers["x-api-key"]).toBe("sk-ant-xyz");
    expect(opts.headers["anthropic-version"]).toBe("2023-06-01");
    expect(opts.headers["anthropic-dangerous-direct-browser-access"]).toBe("true");
    const body = JSON.parse(opts.body);
    expect(body.model).toBe("claude-haiku-4-5");
    expect(body.max_tokens).toBeGreaterThan(0);
    expect(body.messages[0].role).toBe("user");
    expect(body.messages[0].content).toContain("Budget Q2");
  });

  test("wirft SummaryError bei nicht-2xx (Aufrufer faellt zurueck)", async () => {
    mockFetch(401, { error: { message: "invalid key" } });
    await expect(summarizeMail("bad", mail, "x")).rejects.toBeInstanceOf(SummaryError);
  });

  test("wirft SummaryError bei Netzwerkfehler", async () => {
    (global as any).fetch = jest.fn().mockRejectedValue(new TypeError("Failed to fetch"));
    await expect(summarizeMail("k", mail, "x")).rejects.toBeInstanceOf(SummaryError);
  });

  test("wirft SummaryError bei Refusal-Antwort", async () => {
    mockFetch(200, { content: [], stop_reason: "refusal" });
    await expect(summarizeMail("k", mail, "x")).rejects.toBeInstanceOf(SummaryError);
  });
});
