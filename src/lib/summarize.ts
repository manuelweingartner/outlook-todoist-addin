// KI-Zusammenfassung des Mailinhalts als Todoist-Kommentartitel. Client-seitiger
// Direkt-Aufruf an die Anthropic Messages-API (raw fetch, konsistent mit todoist.ts;
// kein SDK -> schlankes Bundle). Der Header anthropic-dangerous-direct-browser-access
// aktiviert CORS fuer Browser-Aufrufe. Fehler werfen SummaryError; der Aufrufer faengt
// und faellt auf den Betreff-Titel zurueck (blockiert nie das Anhaengen).
import { MailData } from "./emlBuilder";

const API_URL = "https://api.anthropic.com/v1/messages";
const MODEL = "claude-haiku-4-5";
const MAX_TOKENS = 120;
const BODY_LIMIT = 2000;

export class SummaryError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "SummaryError";
    Object.setPrototypeOf(this, SummaryError.prototype);
  }
}

export function buildSummaryPrompt(mail: MailData, bodyText: string): string {
  const to = mail.to.length ? mail.to.join(", ") : "(unbekannt)";
  const cc = mail.cc.length ? `\nCC: ${mail.cc.join(", ")}` : "";
  const body = bodyText.slice(0, BODY_LIMIT);
  return [
    "Fasse die folgende E-Mail in GENAU EINEM prägnanten deutschen Satz zusammen:",
    "von wem sie kommt, worum es geht und welche Aktion (falls vorhanden) gewünscht ist.",
    "Nur der Satz, kein Vorspann, keine Anführungszeichen. Echte Umlaute (ä/ö/ü/ß).",
    "",
    `Von: ${mail.from || "(unbekannt)"}`,
    `An: ${to}${cc}`,
    `Betreff: ${mail.subject}`,
    "",
    "Inhalt:",
    body,
  ].join("\n");
}

// Extrahiert den Ein-Satz-Text aus der Messages-API-Antwort. Wirft bei Refusal,
// fehlendem oder leerem Text.
export function parseSummary(json: unknown): string {
  const j = json as { content?: Array<{ type?: string; text?: string }>; stop_reason?: string };
  if (j && j.stop_reason === "refusal") throw new SummaryError("Anthropic hat die Anfrage abgelehnt (refusal).");
  const block = j && Array.isArray(j.content) ? j.content.find((b) => b.type === "text" && typeof b.text === "string") : undefined;
  let text = (block?.text ?? "").trim();
  // Umschliessende Anfuehrungszeichen entfernen (Modell packt den Satz manchmal in "...").
  text = text.replace(/^["'„»]\s*/, "").replace(/\s*["'“«]$/, "").trim();
  if (!text) throw new SummaryError("Leere Zusammenfassung erhalten.");
  return text;
}

export async function summarizeMail(key: string, mail: MailData, bodyText: string): Promise<string> {
  let res: Response;
  try {
    res = await fetch(API_URL, {
      method: "POST",
      headers: {
        "x-api-key": key,
        "anthropic-version": "2023-06-01",
        "content-type": "application/json",
        "anthropic-dangerous-direct-browser-access": "true",
      },
      body: JSON.stringify({
        model: MODEL,
        max_tokens: MAX_TOKENS,
        messages: [{ role: "user", content: buildSummaryPrompt(mail, bodyText) }],
      }),
    });
  } catch (e) {
    throw new SummaryError(`Netzwerkfehler: ${(e as Error).message}`);
  }
  if (!res.ok) {
    const body = await res.text().catch(() => "");
    throw new SummaryError(`Anthropic API ${res.status}: ${body}`);
  }
  return parseSummary(await res.json());
}
