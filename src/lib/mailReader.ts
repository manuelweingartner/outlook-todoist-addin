import { MailData, MailAttachment } from "./emlBuilder";

function fmtAddress(a: { displayName?: string; emailAddress: string }): string {
  return a.displayName ? `${a.displayName} <${a.emailAddress}>` : a.emailAddress;
}

function getHtmlBody(item: any): Promise<string> {
  return new Promise((resolve, reject) => {
    item.body.getAsync(Office.CoercionType.Html, (r: any) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error),
    );
  });
}

function getAttachmentBase64(item: any, id: string): Promise<string | null> {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(id, (r: any) => {
      if (r.status !== Office.AsyncResultStatus.Succeeded) return reject(r.error);
      if (r.value.format === Office.MailboxEnums.AttachmentContentFormat.Base64) resolve(r.value.content);
      else resolve(null);
    });
  });
}

export async function readCurrentMail(): Promise<MailData> {
  const item: any = Office.context.mailbox.item;
  const htmlBody = await getHtmlBody(item);

  const attachments: MailAttachment[] = [];
  const fileAtts = (item.attachments || []).filter(
    (a: any) => a.attachmentType === Office.MailboxEnums.AttachmentType.File,
  );
  for (const a of fileAtts) {
    try {
      const base64 = await getAttachmentBase64(item, a.id);
      if (base64) attachments.push({ name: a.name, contentType: a.contentType, base64 });
    } catch (e) {
      console.error("Anhang nicht lesbar:", a.name, e);
    }
  }

  return {
    subject: item.subject || "(kein Betreff)",
    from: item.from ? fmtAddress(item.from) : "",
    to: (item.to || []).map(fmtAddress),
    cc: (item.cc || []).map(fmtAddress),
    date: (item.dateTimeCreated ? new Date(item.dateTimeCreated) : new Date()).toUTCString(),
    htmlBody,
    attachments,
  };
}
