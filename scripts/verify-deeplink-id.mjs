// Verifiziert die Deep-Link-Id-Frage: Liefert /api/v1/id_mappings fuer eine
// echte (neue, alphanumerische) Task-Id die alte numerische Id? Und gibt beide
// aus, damit man todoist://task?id=<numerisch> vs. <alphanumerisch> im Desktop
// testen kann.
//
// Token NUR ueber stdin (nie als Argument/Chat):
//   node scripts/verify-deeplink-id.mjs
//   <Token einfuegen, Enter, Ctrl+Z Enter (Windows) / Ctrl+D (Unix)>
// Oder:  echo %TODOIST_TOKEN% | node scripts/verify-deeplink-id.mjs   (nicht empfohlen: Shell-History)

const API = "https://api.todoist.com/api/v1";

async function readStdin() {
  const chunks = [];
  for await (const c of process.stdin) chunks.push(c);
  return Buffer.concat(chunks).toString("utf8").trim();
}

async function main() {
  const token = process.env.TODOIST_TOKEN || (await readStdin());
  if (!token) { console.error("Kein Token (stdin leer, TODOIST_TOKEN nicht gesetzt)."); process.exit(1); }
  const h = { Authorization: `Bearer ${token}` };

  const res = await fetch(`${API}/tasks?limit=5`, { headers: h });
  if (!res.ok) { console.error("Task-Abruf fehlgeschlagen:", res.status, await res.text()); process.exit(1); }
  const data = await res.json();
  const tasks = Array.isArray(data) ? data : (data.results ?? []);
  if (tasks.length === 0) { console.error("Keine offenen Tasks gefunden."); process.exit(1); }

  const ids = tasks.map((t) => t.id);
  const map = await fetch(`${API}/id_mappings/tasks/${ids.join(",")}`, { headers: h });
  const mappings = map.ok ? await map.json() : [];
  const byNew = Object.fromEntries(mappings.map((m) => [m.new_id, m.old_id]));

  console.log("id_mappings HTTP:", map.status);
  console.log("");
  for (const t of tasks) {
    const oldId = byNew[t.id] ?? "(kein Mapping)";
    console.log(`Task: ${JSON.stringify(t.content).slice(0, 50)}`);
    console.log(`  neu (v1/API):     ${t.id}`);
    console.log(`  alt (numerisch):  ${oldId}`);
    console.log(`  Test-Desktop:     todoist://task?id=${oldId}`);
    console.log(`  Test-Desktop-alt: todoist://task?id=${t.id}`);
    console.log("");
  }
  console.log("Naechster Schritt: beide todoist://-Links im Desktop-Client feuern und schauen, welcher zur Task springt.");
}

main().catch((e) => { console.error("FEHLER:", e.message); process.exit(1); });
