import { formatMailDate, htmlToText, prepareCurrentMail } from "../src/lib/attachToTask";
import { readCurrentMail } from "../src/lib/mailReader";

jest.mock("../src/lib/mailReader");

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
