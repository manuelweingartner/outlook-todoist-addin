# Design: KI-Zusammenfassung als Todoist-Kommentartitel

Datum: 2026-07-07

## Ziel

Statt `Betreff (Datum, von Name)` soll der Todoist-Kommentar beim Anhängen einer Mail
einen **einzelnen, prägnanten deutschen Satz** als Titel tragen, der aus dem Mailinhalt
zusammenfasst: von wem, worum es geht, gewünschte Aktion. Erzeugt client-seitig über
einen persönlichen Anthropic-Key.

## Verhalten

- Der Anhäng-Flow bleibt unverändert schnell: Klick auf Task -> sofort anhängen ->
  bestehender Rückgängig-Button. Kein Vorschau-/Bestätigungsschritt.
- Neu: der Kommentartitel ist der KI-Satz **statt** `Betreff (Datum, von Name)`. Der
  exakte Betreff bleibt in der angehängten `.eml` erhalten (Wiederfinden nicht verloren).

## Feature-Gating (nicht-brechend)

- Aktiv **genau dann, wenn ein Anthropic-Key hinterlegt ist**. Kein Key -> Verhalten wie
  heute (`buildCommentText`, Betreff-Titel). Key eintragen = einschalten, löschen = aus.
- Kein separater An/Aus-Schalter (YAGNI): Vorhandensein des Keys ist der Schalter.

## Key-Eingabe / Speicherung

- Zweites Feld im Einstellungs-/Onboarding-Bereich, analog zum Todoist-Token.
- Speicherung in `Office.context.roamingSettings` unter neuem Schlüssel `anthropicKey`
  (bestehender Todoist-Key bleibt `todoistToken`).
- Datenschutz-Hinweis direkt am Feld: "Für die Zusammenfassung wird der Mailinhalt an
  Anthropic gesendet." Plus Transparenz: Key liegt (wie der Todoist-Token) im Browser/
  roamingSettings (technisch auslesbar) und roamt über das Exchange-Postfach.

## Modell & Aufruf

- Claude **Haiku 4.5** (`claude-haiku-4-5`), ein einzelner nicht-streamender Aufruf.
- Direkt-Browser-Call an `https://api.anthropic.com/v1/messages`, raw `fetch` (konsistent
  mit `todoist.ts`, das ebenfalls raw fetch nutzt; kein SDK -> schlankes Bundle).
- Header: `x-api-key: <key>`, `anthropic-version: 2023-06-01`, `content-type: application/json`,
  `anthropic-dangerous-direct-browser-access: true` (aktiviert CORS für Browser-Aufrufe).
- Body: `{ model: "claude-haiku-4-5", max_tokens: ~100, messages: [{role:"user", content: <prompt>}] }`.
- ~1s Latenz, Bruchteile eines Rappens pro Mail.

## Prompt

Input: Absender (`from`), Empfänger (`to`/`cc` sofern via Office.js verfügbar), Betreff,
und der bereits extrahierte `bodyText` (max 2000 Zeichen, aus `prepareMail`).
Auftrag: **genau ein prägnanter deutscher Satz** — von wem, worum es geht, gewünschte
Aktion. Nur der Satz, kein Vorspann, keine Anführungszeichen. Echte Umlaute.

## Fehler-Fallback (blockiert nie das Anhängen)

- Key fehlt, API-Fehler (nicht-2xx), `stop_reason: "refusal"`, leere/keine Textantwort,
  oder Netzwerkfehler -> **zurück auf den heutigen Betreff-Titel** (`buildCommentText`).
- Nicht-blockierender Status-Hinweis: "Zusammenfassung fehlgeschlagen, Betreff verwendet."
- Fehler wird per `console.error` geloggt (kein stilles `try/catch`, Repo-Regel).

## Module & Datenfluss

- **Neu `src/lib/settings.ts`:** `getAnthropicKey()` / `setAnthropicKey()` analog
  `getToken`/`setToken` (gleiche `saveAsync`-Mechanik, Schlüssel `anthropicKey`).
- **Neu `src/lib/summarize.ts`:**
  - `buildSummaryPrompt(mail: MailData, bodyText: string): string` — rein, testbar.
  - `parseSummary(apiJson: unknown): string` — extrahiert `content[0].text`, trimmt,
    entfernt umschließende Anführungszeichen; wirft bei leer/Refusal.
  - `summarizeMail(key, mail, bodyText): Promise<string>` — fetch + parse; wirft bei
    Fehler (Aufrufer fängt und fällt zurück).
- **`src/lib/attachToTask.ts`:** Der Aufrufer (Attach-Flow in `taskpane.ts`, wo heute
  `prepared.commentText` verwendet wird) bekommt eine Titel-Auflösung:
  wenn Anthropic-Key vorhanden -> `summarizeMail(...)` versuchen -> Ergebnis als Titel;
  sonst / bei Fehler -> `prepared.commentText` (heutiger Betreff-Titel).
  Umsetzung: `attachPreparedToTask` erhält optional den Titel-Text übergeben, oder eine
  kleine `resolveCommentTitle(mailData, prepared)`-Funktion kapselt die Key-Prüfung +
  Fallback. Genaue Signatur im Plan.
- **`src/taskpane/taskpane.ts` + `.html`:** zweites Key-Feld im Einstellungsbereich,
  Datenschutz-Hinweis, Speichern via `setAnthropicKey`. Attach-Status zeigt bei Fallback
  den Hinweis.

## Tests (TDD)

- `summarize.test.ts`: Prompt-Bau (enthält Absender/Betreff/Body-Ausschnitt), Antwort-
  Parsing (Text-Extraktion, Quote-Trim), Fehlerfälle (nicht-2xx wirft, Refusal wirft,
  leere Antwort wirft) mit gemocktem `fetch` — analog `todoist.test.ts`.
- `settings.test.ts`: `getAnthropicKey`/`setAnthropicKey` analog Todoist-Token.
- Fallback-Verhalten der Titel-Auflösung (Key fehlt -> Betreff-Titel; summarize wirft ->
  Betreff-Titel) als Unit-Test.

## Nicht im Scope (YAGNI)

- Kein separater Toggle, keine Modellauswahl-UI, keine Vorschau/Editier-UI, kein Caching,
  kein Streaming, kein Backend/Proxy. Betreff + Absender/Empfänger als reines
  Metadaten-Format (ohne KI) wurde verworfen zugunsten "nur Zusammenfassung".
