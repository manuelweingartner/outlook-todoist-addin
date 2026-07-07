import { getToken, setToken, getAnthropicKey, setAnthropicKey } from "../src/lib/settings";

describe("settings", () => {
  let store: Record<string, unknown>;
  beforeEach(() => {
    store = {};
    (global as any).Office.context.roamingSettings = {
      get: (k: string) => store[k],
      set: (k: string, v: unknown) => { store[k] = v; },
      saveAsync: (cb: (r: any) => void) =>
        cb({ status: (global as any).Office.AsyncResultStatus.Succeeded }),
    };
  });

  test("getToken liefert null wenn nichts gespeichert", () => {
    expect(getToken()).toBeNull();
  });

  test("setToken speichert, getToken liest zurueck", async () => {
    await setToken("abc123");
    expect(getToken()).toBe("abc123");
  });

  test("setToken lehnt ab wenn saveAsync fehlschlaegt", async () => {
    (global as any).Office.context.roamingSettings.saveAsync = (cb: (r: any) => void) =>
      cb({ status: "failed", error: new Error("save error") });
    await expect(setToken("x")).rejects.toBeInstanceOf(Error);
  });

  test("getToken liefert null nach setToken('')", async () => {
    await setToken("");
    expect(getToken()).toBeNull();
  });

  test("Anthropic-Key: null wenn nichts gespeichert", () => {
    expect(getAnthropicKey()).toBeNull();
  });

  test("Anthropic-Key: setzen und zurueticklesen, getrennt vom Todoist-Token", async () => {
    await setToken("todoist-tok");
    await setAnthropicKey("sk-ant-xyz");
    expect(getAnthropicKey()).toBe("sk-ant-xyz");
    expect(getToken()).toBe("todoist-tok"); // Keys beeinflussen sich nicht
  });

  test("Anthropic-Key: leerer String loescht (getAnthropicKey -> null)", async () => {
    await setAnthropicKey("sk-ant-xyz");
    await setAnthropicKey("");
    expect(getAnthropicKey()).toBeNull();
  });
});
