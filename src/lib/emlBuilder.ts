export interface MailAttachment { name: string; contentType: string; base64: string; }
export interface MailData {
  subject: string;
  from: string;
  to: string[];
  cc: string[];
  date: string;
  htmlBody: string;
  attachments: MailAttachment[];
}

const CRLF = "\r\n";

export function utf8ToBase64(str: string): string {
  return btoa(unescape(encodeURIComponent(str)));
}

export function encodeHeader(value: string): string {
  if (/^[\x20-\x7E]*$/.test(value)) return value;
  return `=?UTF-8?B?${utf8ToBase64(value)}?=`;
}

function chunk76(b64: string): string {
  return b64.replace(/(.{76})/g, `$1${CRLF}`);
}

function htmlPart(html: string): string {
  return (
    `Content-Type: text/html; charset=UTF-8${CRLF}` +
    `Content-Transfer-Encoding: base64${CRLF}${CRLF}` +
    chunk76(utf8ToBase64(html))
  );
}

export function buildEml(mail: MailData, boundary = "=_todoist_eml_boundary_"): string {
  const headers: string[] = [
    `Subject: ${encodeHeader(mail.subject)}`,
    `From: ${mail.from}`,
  ];
  if (mail.to.length) headers.push(`To: ${mail.to.join(", ")}`);
  if (mail.cc.length) headers.push(`Cc: ${mail.cc.join(", ")}`);
  headers.push(`Date: ${mail.date}`);
  headers.push(`MIME-Version: 1.0`);

  if (mail.attachments.length === 0) {
    return headers.join(CRLF) + CRLF + htmlPart(mail.htmlBody) + CRLF;
  }

  headers.push(`Content-Type: multipart/mixed; boundary="${boundary}"`);
  const parts: string[] = [`--${boundary}${CRLF}${htmlPart(mail.htmlBody)}`];
  for (const att of mail.attachments) {
    parts.push(
      `--${boundary}${CRLF}` +
        `Content-Type: ${att.contentType}; name="${att.name}"${CRLF}` +
        `Content-Transfer-Encoding: base64${CRLF}` +
        `Content-Disposition: attachment; filename="${att.name}"${CRLF}${CRLF}` +
        chunk76(att.base64),
    );
  }
  return headers.join(CRLF) + CRLF + CRLF + parts.join(CRLF) + CRLF + `--${boundary}--${CRLF}`;
}
