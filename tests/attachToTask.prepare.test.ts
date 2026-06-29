import { formatMailDate } from "../src/lib/attachToTask";

describe("formatMailDate", () => {
  test("formatiert UTC-String zu dd.mm.yyyy", () => {
    expect(formatMailDate("Mon, 05 Jan 2026 10:00:00 +0000")).toBe("05.01.2026");
  });
  test("leerer/ungueltiger String gibt leer", () => {
    expect(formatMailDate("")).toBe("");
    expect(formatMailDate("kein datum")).toBe("");
  });
  test("kurz nach Mitternacht UTC bleibt am gleichen Tag (kein Timezone-Drift)", () => {
    expect(formatMailDate("Mon, 05 Jan 2026 00:30:00 +0000")).toBe("05.01.2026");
  });
});
