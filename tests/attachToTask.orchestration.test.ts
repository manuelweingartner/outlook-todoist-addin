jest.mock("../src/lib/mailReader");
jest.mock("../src/lib/todoist");

import { attachCurrentMailToTask, attachPreparedToTask, prepareCurrentMail, MAX_BYTES } from "../src/lib/attachToTask";
import { readCurrentMail } from "../src/lib/mailReader";
import { uploadFile, addComment } from "../src/lib/todoist";
import { MailData } from "../src/lib/emlBuilder";

const baseMail: MailData = {
  subject: "S",
  from: "a@b.ch",
  to: [],
  cc: [],
  date: "Mon, 01 Jan 2026 10:00:00 +0000",
  htmlBody: "<p>x</p>",
  attachments: [],
};

describe("attachCurrentMailToTask", () => {
  beforeEach(() => jest.clearAllMocks());

  test("happy path: Kommentartext = Betreff (Datum, von Name), gibt id zurueck", async () => {
    (readCurrentMail as jest.Mock).mockResolvedValue(baseMail);
    (uploadFile as jest.Mock).mockResolvedValue({ file_url: "u", file_name: "S.eml", file_type: "message/rfc822" });
    (addComment as jest.Mock).mockResolvedValue("c1");

    const id = await attachCurrentMailToTask("tok", "task1");

    expect(id).toBe("c1");
    const [tok, taskId, file, content] = (addComment as jest.Mock).mock.calls[0];
    expect(tok).toBe("tok");
    expect(taskId).toBe("task1");
    expect(file.file_url).toBe("u");
    expect(content).toBe("S (01.01.2026, von a@b.ch)");
  });

  test("prepareCurrentMail liefert sizeBytes + commentText", async () => {
    (readCurrentMail as jest.Mock).mockResolvedValue(baseMail);
    const p = await prepareCurrentMail();
    expect(p.subject).toBe("S");
    expect(p.commentText).toBe("S (01.01.2026, von a@b.ch)");
    expect(p.sizeBytes).toBeGreaterThan(0);
  });

  test("wirft bei >25 MB und laedt NICHT hoch", async () => {
    const huge = "x".repeat(MAX_BYTES);
    (readCurrentMail as jest.Mock).mockResolvedValue({
      ...baseMail,
      htmlBody: huge,
    });

    await expect(attachCurrentMailToTask("tok", "task1")).rejects.toThrow(
      /25 MB/
    );
    expect(uploadFile).not.toHaveBeenCalled();
    expect(addComment).not.toHaveBeenCalled();
  });
});
