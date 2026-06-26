import { buildEml, encodeHeader, utf8ToBase64, MailData } from "../src/lib/emlBuilder";

const base: MailData = {
  subject: "Hallo",
  from: "A B <a@b.ch>",
  to: ["C D <c@d.ch>"],
  cc: [],
  date: "Mon, 01 Jan 2026 10:00:00 +0000",
  htmlBody: "<p>Text</p>",
  attachments: [],
};

describe("encodeHeader", () => {
  test("ASCII bleibt unveraendert", () => {
    expect(encodeHeader("Hello")).toBe("Hello");
  });
  test("Umlaute werden RFC2047-kodiert", () => {
    const out = encodeHeader("Grüezi");
    expect(out.startsWith("=?UTF-8?B?")).toBe(true);
    expect(out.endsWith("?=")).toBe(true);
  });
  test("Werte mit Steuerzeichen werden kodiert (kein Header-Injection)", () => {
    const out = encodeHeader("a\r\nInjected: x");
    expect(out.startsWith("=?UTF-8?B?")).toBe(true);
    expect(out.endsWith("?=")).toBe(true);
  });
});

describe("buildEml ohne Anhang", () => {
  const eml = buildEml(base);
  test("enthaelt Header", () => {
    expect(eml).toContain("Subject: Hallo");
    expect(eml).toContain("From: A B <a@b.ch>");
    expect(eml).toContain("To: C D <c@d.ch>");
    expect(eml).toContain("MIME-Version: 1.0");
  });
  test("Body ist base64-kodiert", () => {
    expect(eml).toContain("Content-Transfer-Encoding: base64");
    expect(eml).toContain(utf8ToBase64("<p>Text</p>"));
  });
  test("single-part: Header-Block durch Leerzeile vom Body getrennt, Body dekodiert zu htmlBody", () => {
    const idx = eml.indexOf("\r\n\r\n");
    expect(idx).toBeGreaterThan(0);
    const headerBlock = eml.slice(0, idx);
    const body = eml.slice(idx + 4).trim();
    expect(headerBlock).toContain("Content-Type: text/html; charset=UTF-8");
    expect(headerBlock).toContain("MIME-Version: 1.0");
    expect(Buffer.from(body, "base64").toString("utf8")).toBe("<p>Text</p>");
  });
});

describe("buildEml mit Anhang", () => {
  const eml = buildEml(
    { ...base, attachments: [{ name: "doc.pdf", contentType: "application/pdf", base64: "QUJD" }] },
    "BOUND",
  );
  test("ist multipart/mixed mit Boundary", () => {
    expect(eml).toContain('Content-Type: multipart/mixed; boundary="BOUND"');
    expect(eml).toContain("--BOUND");
    expect(eml).toContain("--BOUND--");
  });
  test("enthaelt Anhang-Teil", () => {
    expect(eml).toContain('Content-Disposition: attachment; filename="doc.pdf"');
    expect(eml).toContain("QUJD");
  });
});
