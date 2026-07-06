import { buildEml, MailData } from "./emlBuilder";
import { readCurrentMail } from "./mailReader";
import { uploadFile, addComment } from "./todoist";

export const MAX_BYTES = 25 * 1024 * 1024;

export interface PreparedMail {
  blob: Blob;
  fileName: string;
  sizeBytes: number;
  subject: string;
  commentText: string;
  bodyText: string;
}

function sanitizeFileName(subject: string): string {
  const cleaned = (subject || "Mail").replace(/[\\/:*?"<>|]/g, "_").trim();
  return `${cleaned.slice(0, 120) || "Mail"}.eml`;
}

// Regex-Strip reicht fuer Scoring-Zwecke (kein DOM noetig, laeuft auch in Tests ohne jsdom-Kosten).
export function htmlToText(html: string): string {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&quot;/gi, '"')
    .replace(/&#\d+;/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

export function emlBlobFor(mail: MailData): { blob: Blob; fileName: string } {
  const eml = buildEml(mail);
  const blob = new Blob([eml], { type: "message/rfc822" });
  return { blob, fileName: sanitizeFileName(mail.subject) };
}

export function formatMailDate(utc: string): string {
  const d = new Date(utc);
  if (isNaN(d.getTime())) return "";
  const dd = String(d.getUTCDate()).padStart(2, "0");
  const mm = String(d.getUTCMonth() + 1).padStart(2, "0");
  return `${dd}.${mm}.${d.getUTCFullYear()}`;
}

export async function prepareCurrentMail(): Promise<PreparedMail> {
  const mail = await readCurrentMail();
  const { blob, fileName } = emlBlobFor(mail);
  const datePart = formatMailDate(mail.date);
  const commentText = datePart ? `${mail.subject} (${datePart})` : mail.subject;
  const bodyText = htmlToText(mail.htmlBody).slice(0, 2000);
  return { blob, fileName, sizeBytes: blob.size, subject: mail.subject, commentText, bodyText };
}

export async function attachPreparedToTask(token: string, taskId: string, prepared: PreparedMail): Promise<string> {
  if (prepared.sizeBytes > MAX_BYTES) {
    throw new Error(`Mail ist ${(prepared.sizeBytes / 1024 / 1024).toFixed(1)} MB gross, Todoist erlaubt max 25 MB.`);
  }
  const uploaded = await uploadFile(token, prepared.blob, prepared.fileName);
  return addComment(token, taskId, uploaded, prepared.commentText);
}

export async function attachCurrentMailToTask(token: string, taskId: string): Promise<string> {
  const prepared = await prepareCurrentMail();
  return attachPreparedToTask(token, taskId, prepared);
}
