const KEY = "todoistToken";
const ANTHROPIC_KEY = "anthropicKey";

function readSetting(key: string): string | null {
  const v = Office.context.roamingSettings.get(key) as string | undefined;
  return v && v.length > 0 ? v : null;
}

function writeSetting(key: string, value: string): Promise<void> {
  Office.context.roamingSettings.set(key, value);
  return new Promise<void>((resolve, reject) => {
    Office.context.roamingSettings.saveAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}

export function getToken(): string | null {
  return readSetting(KEY);
}

export function setToken(token: string): Promise<void> {
  return writeSetting(KEY, token);
}

// Optionaler Anthropic-Key fuer die KI-Zusammenfassung des Kommentartitels.
// Vorhandensein = Feature an; leer/null = aus (Fallback auf Betreff-Titel).
export function getAnthropicKey(): string | null {
  return readSetting(ANTHROPIC_KEY);
}

export function setAnthropicKey(key: string): Promise<void> {
  return writeSetting(ANTHROPIC_KEY, key);
}
