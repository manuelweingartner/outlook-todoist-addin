jest.mock("../src/lib/mailReader");
jest.mock("../src/lib/todoist");

import { attachCurrentMailToTask, MAX_BYTES } from "../src/lib/attachToTask";
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

  test("happy path: liest Mail, laedt hoch, kommentiert in Reihenfolge", async () => {
    (readCurrentMail as jest.Mock).mockResolvedValue(baseMail);
    (uploadFile as jest.Mock).mockResolvedValue({
      file_url: "u",
      file_name: "S.eml",
      file_type: "message/rfc822",
    });
    (addComment as jest.Mock).mockResolvedValue(undefined);

    await attachCurrentMailToTask("tok", "task1");

    expect(readCurrentMail).toHaveBeenCalled();
    expect(uploadFile).toHaveBeenCalledTimes(1);
    const [tok, , fileName] = (uploadFile as jest.Mock).mock.calls[0];
    expect(tok).toBe("tok");
    expect(fileName).toBe("S.eml");
    expect(addComment).toHaveBeenCalledWith("tok", "task1", {
      file_url: "u",
      file_name: "S.eml",
      file_type: "message/rfc822",
    });
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
