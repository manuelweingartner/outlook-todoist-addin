import { buildEml, MailData } from "./emlBuilder";
import { readCurrentMail } from "./mailReader";
import { uploadFile, addComment } from "./todoist";

export const MAX_BYTES = 25 * 1024 * 1024;

function sanitizeFileName(subject: string): string {
  const cleaned = (subject || "Mail").replace(/[\\/:*?"<>|]/g, "_").trim();
  return `${cleaned.slice(0, 120) || "Mail"}.eml`;
}

export function emlBlobFor(mail: MailData): { blob: Blob; fileName: string } {
  const eml = buildEml(mail);
  const blob = new Blob([eml], { type: "message/rfc822" });
  return { blob, fileName: sanitizeFileName(mail.subject) };
}

export async function attachCurrentMailToTask(token: string, taskId: string): Promise<void> {
  const mail = await readCurrentMail();
  const { blob, fileName } = emlBlobFor(mail);
  if (blob.size > MAX_BYTES) {
    throw new Error(`Mail ist ${(blob.size / 1024 / 1024).toFixed(1)} MB gross, Todoist erlaubt max 25 MB.`);
  }
  const uploaded = await uploadFile(token, blob, fileName);
  await addComment(token, taskId, uploaded);
}
