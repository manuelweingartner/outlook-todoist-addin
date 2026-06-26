const KEY = "todoistToken";

export function getToken(): string | null {
  const v = Office.context.roamingSettings.get(KEY) as string | undefined;
  return v && v.length > 0 ? v : null;
}

export function setToken(token: string): Promise<void> {
  Office.context.roamingSettings.set(KEY, token);
  return new Promise<void>((resolve, reject) => {
    Office.context.roamingSettings.saveAsync((res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(res.error);
    });
  });
}
