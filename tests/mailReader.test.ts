import { readCurrentMail } from "../src/lib/mailReader";

function mockItem() {
  const O = (global as any).Office;
  (global as any).Office.context.mailbox = {
    item: {
      subject: "Betreff",
      from: { displayName: "Sender", emailAddress: "s@x.ch" },
      to: [{ displayName: "Empf", emailAddress: "e@x.ch" }],
      cc: [],
      dateTimeCreated: new Date("2026-01-01T10:00:00Z"),
      body: {
        getAsync: (_t: string, cb: (r: any) => void) =>
          cb({ status: O.AsyncResultStatus.Succeeded, value: "<p>Hi</p>" }),
      },
      attachments: [
        { id: "1", name: "a.pdf", contentType: "application/pdf", attachmentType: "file" },
        { id: "2", name: "inline.png", contentType: "image/png", attachmentType: "item" },
      ],
      getAttachmentContentAsync: (id: string, cb: (r: any) => void) =>
        cb({
          status: O.AsyncResultStatus.Succeeded,
          value: { format: O.MailboxEnums.AttachmentContentFormat.Base64, content: "B64-" + id },
        }),
    },
  };
}

test("liest Betreff, Adressen, Body und nur File-Anhaenge", async () => {
  mockItem();
  const mail = await readCurrentMail();
  expect(mail.subject).toBe("Betreff");
  expect(mail.from).toBe("Sender <s@x.ch>");
  expect(mail.to).toEqual(["Empf <e@x.ch>"]);
  expect(mail.htmlBody).toBe("<p>Hi</p>");
  expect(mail.attachments).toHaveLength(1);
  expect(mail.attachments[0]).toEqual({ name: "a.pdf", contentType: "application/pdf", base64: "B64-1" });
});

test("Anhang-Lesefehler wird geloggt und uebersprungen, Rest bleibt", async () => {
  const O = (global as any).Office;
  const errSpy = jest.spyOn(console, "error").mockImplementation(() => {});
  (global as any).Office.context.mailbox = {
    item: {
      subject: "S", from: { emailAddress: "a@x.ch" }, to: [], cc: [],
      dateTimeCreated: new Date("2026-01-01T10:00:00Z"),
      body: { getAsync: (_t: string, cb: (r: any) => void) => cb({ status: O.AsyncResultStatus.Succeeded, value: "<p>x</p>" }) },
      attachments: [
        { id: "1", name: "bad.pdf", contentType: "application/pdf", attachmentType: "file" },
        { id: "2", name: "good.pdf", contentType: "application/pdf", attachmentType: "file" },
      ],
      getAttachmentContentAsync: (id: string, cb: (r: any) => void) => {
        if (id === "1") cb({ status: O.AsyncResultStatus.Failed, error: new Error("nope") });
        else cb({ status: O.AsyncResultStatus.Succeeded, value: { format: O.MailboxEnums.AttachmentContentFormat.Base64, content: "B64-" + id } });
      },
    },
  };
  const mail = await readCurrentMail();
  expect(mail.attachments).toHaveLength(1);
  expect(mail.attachments[0].name).toBe("good.pdf");
  expect(errSpy).toHaveBeenCalled();
  errSpy.mockRestore();
});
