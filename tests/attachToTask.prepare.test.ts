import { formatMailDate, htmlToText, prepareCurrentMail, senderName, prepareMail } from "../src/lib/attachToTask";
import { readCurrentMail } from "../src/lib/mailReader";
import { MailData } from "../src/lib/emlBuilder";

jest.mock("../src/lib/mailReader");

const mailFixture = (over: Partial<MailData> = {}): MailData => ({
  subject: "Rechnung Q3",
  from: "Joel Willi <joel@wowoni.ch>",
  to: [],
  cc: [],
  date: "Sun, 05 Jul 2026 10:00:00 GMT",
  htmlBody: "<p>Hallo</p>",
  attachments: [{ name: "a.pdf", contentType: "application/pdf", base64: "QUJD" }],
  ...over,
});

describe("htmlToText", () => {
  test("strippt Tags und kollabiert Whitespace", () => {
    expect(htmlToText("<div><p>Hallo  <b>Welt</b></p>\n<p>Zeile 2</p></div>")).toBe("Hallo Welt Zeile 2");
  });

  test("entfernt style- und script-Bloecke samt Inhalt", () => {
    expect(htmlToText("<style>p{color:red}</style><script>var x=1;</script>Text")).toBe("Text");
  });

  test("dekodiert gaengige Entities", () => {
    expect(htmlToText("A&nbsp;&amp;&nbsp;B &lt;C&gt; &quot;D&quot;")).toBe('A & B <C> "D"');
  });
});

describe("prepareCurrentMail bodyText", () => {
  test("liefert gestrippten Body, gekappt auf 2000 Zeichen", async () => {
    const mockMailData = {
      subject: "Test",
      from: "test@example.com",
      to: ["recipient@example.com"],
      cc: [],
      date: "Mon, 05 Jan 2026 10:00:00 +0000",
      htmlBody: "<p>" + "a".repeat(3000) + "</p>",
      attachments: [],
    };
    (readCurrentMail as jest.Mock).mockResolvedValue(mockMailData);
    const prepared = await prepareCurrentMail();
    expect(prepared.bodyText).toHaveLength(2000);
    expect(prepared.bodyText.startsWith("aaa")).toBe(true);
  });
});

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

describe("senderName", () => {
  test("Name <adresse> liefert den Namen", () => {
    expect(senderName("Joel Willi <joel@wowoni.ch>")).toBe("Joel Willi");
  });
  test("nur Adresse liefert die Adresse", () => {
    expect(senderName("joel@wowoni.ch")).toBe("joel@wowoni.ch");
  });
  test("leer liefert leer", () => {
    expect(senderName("")).toBe("");
  });
});

describe("prepareMail commentText", () => {
  test("Betreff (Datum, von Name)", () => {
    expect(prepareMail(mailFixture(), true).commentText).toBe("Rechnung Q3 (05.07.2026, von Joel Willi)");
  });
  test("ohne Absender: Betreff (Datum)", () => {
    expect(prepareMail(mailFixture({ from: "" }), true).commentText).toBe("Rechnung Q3 (05.07.2026)");
  });
  test("ohne Datum: Betreff (von Name)", () => {
    expect(prepareMail(mailFixture({ date: "kaputt" }), true).commentText).toBe("Rechnung Q3 (von Joel Willi)");
  });
  test("ohne beides: nur Betreff", () => {
    expect(prepareMail(mailFixture({ from: "", date: "kaputt" }), true).commentText).toBe("Rechnung Q3");
  });
});

describe("prepareMail ohne Anhaenge", () => {
  test("laesst Datei-Anhaenge weg und wird kleiner", async () => {
    const full = prepareMail(mailFixture(), true);
    const textOnly = prepareMail(mailFixture(), false);
    expect(textOnly.sizeBytes).toBeLessThan(full.sizeBytes);
    const eml = await new Promise<string>((resolve) => {
      const reader = new FileReader();
      reader.onload = () => resolve(reader.result as string);
      reader.readAsText(textOnly.blob);
    });
    expect(eml).not.toContain("a.pdf");
    expect(eml).toContain("Rechnung Q3");
  });
});
