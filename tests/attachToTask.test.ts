import { emlBlobFor, MAX_BYTES } from "../src/lib/attachToTask";
import { MailData } from "../src/lib/emlBuilder";

const mail: MailData = {
  subject: "Re: Angebot / Offerte",
  from: "a@b.ch", to: [], cc: [], date: "Mon, 01 Jan 2026 10:00:00 +0000",
  htmlBody: "<p>x</p>", attachments: [],
};

test("MAX_BYTES ist 25 MB", () => {
  expect(MAX_BYTES).toBe(25 * 1024 * 1024);
});

test("emlBlobFor erzeugt .eml-Dateinamen aus bereinigtem Betreff", () => {
  const { blob, fileName } = emlBlobFor(mail);
  expect(fileName).toBe("Re_ Angebot _ Offerte.eml");
  expect(blob.type).toBe("message/rfc822");
});

test("emlBlobFor faellt auf Mail.eml zurueck bei leerem Betreff", () => {
  const { fileName } = emlBlobFor({ ...mail, subject: "" });
  expect(fileName).toBe("Mail.eml");
});
