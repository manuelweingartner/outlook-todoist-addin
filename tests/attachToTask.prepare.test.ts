import { formatMailDate } from "../src/lib/attachToTask";

describe("formatMailDate", () => {
  test("formatiert UTC-String zu dd.mm.yyyy", () => {
    expect(formatMailDate("Mon, 05 Jan 2026 10:00:00 +0000")).toBe("05.01.2026");
  });
  test("leerer/ungueltiger String gibt leer", () => {
    expect(formatMailDate("")).toBe("");
    expect(formatMailDate("kein datum")).toBe("");
  });
});
