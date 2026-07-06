import { TodoistTask, NewTaskOptions } from "../lib/todoist";

export function todayIso(now: Date): string {
  const mm = String(now.getMonth() + 1).padStart(2, "0");
  const dd = String(now.getDate()).padStart(2, "0");
  return `${now.getFullYear()}-${mm}-${dd}`;
}

export interface GroupedTasks { overdue: TodoistTask[]; today: TodoistTask[]; }

export function groupTasks(tasks: TodoistTask[], today: string): GroupedTasks {
  const overdue: TodoistTask[] = [];
  const todayList: TodoistTask[] = [];
  for (const task of tasks) {
    const date = task.due?.date;
    if (date && date < today) overdue.push(task);
    else todayList.push(task);
  }
  return { overdue, today: todayList };
}

export function priorityColor(priority?: number): string {
  switch (priority) {
    case 4: return "#e44332"; // P1 rot
    case 3: return "#eb8909"; // P2 orange
    case 2: return "#246fe0"; // P3 blau
    default: return "#808080"; // P4 / keine
  }
}

// Die Outlook-Webview blockt eigene Protokolle (todoist:// zeigt eine
// Verbotssymbol-Fehlerseite). Deshalb zeigt der Link relativ auf unsere
// Umleitungsseite: window.open oeffnet sie im System-Browser, und DER darf
// todoist:// ausloesen (Desktop-Client), mit Web-App als Fallback.
// Die Id ist die neue alphanumerische v1-Id; genau die oeffnet todoist://task?id=
// die korrekte Task (verifiziert). KEINE Uebersetzung auf die alte numerische Id -
// die lehnt der Desktop-Client ab.
export function taskDeepLink(id: string): string {
  return `open-task.html?id=${encodeURIComponent(id)}`;
}

// Client-seitige Suche: case-insensitiv, alle Woerter muessen matchen (UND).
// Woerter mit fuehrendem # matchen NUR den Projektnamen ("#sap rechnung"),
// Woerter ohne # matchen Task-Titel ODER Projektname. Nacktes # wird ignoriert.
export function filterTasks(
  tasks: TodoistTask[],
  query: string,
  projectNames: Record<string, string>,
): TodoistTask[] {
  const words = query.toLowerCase().split(/\s+/).filter(Boolean);
  if (words.length === 0) return tasks;
  return tasks.filter((task) => {
    const project = (projectNames[task.project_id] ?? "").toLowerCase();
    const hay = `${task.content.toLowerCase()} ${project}`;
    return words.every((w) => {
      if (w.startsWith("#")) {
        const p = w.slice(1);
        return p === "" ? true : project.includes(p);
      }
      return hay.includes(w);
    });
  });
}

// Repliziert den frueheren Server-Filter "(today | overdue)" client-seitig.
export function dueTodayOrOverdue(tasks: TodoistTask[], today: string): TodoistTask[] {
  return tasks.filter((task) => !!task.due?.date && task.due.date <= today);
}

// Kleine DE/EN-Stoppwortliste fuer das Vorschlags-Scoring. Bewusst kurz:
// falsche Positives kosten nur einen schwachen Vorschlag, keine Fehler.
const STOPWORDS = new Set([
  "der", "die", "das", "den", "dem", "des", "ein", "eine", "einen", "einem", "einer",
  "und", "oder", "aber", "nicht", "mit", "von", "für", "auf", "aus", "bei", "nach",
  "über", "unter", "vor", "zum", "zur", "ist", "sind", "war", "wird", "werden",
  "wurde", "hat", "haben", "kann", "können", "muss", "müssen", "soll", "sich",
  "auch", "noch", "nur", "schon", "wie", "wir", "ihr", "sie", "ich",
  "the", "and", "for", "are", "but", "not", "with", "from", "this", "that", "you",
  "your", "have", "has", "was", "were", "will", "would", "can", "could", "should",
  "our", "all", "any", "been", "they", "them", "their",
  "hallo", "liebe", "lieber", "grüsse", "gruss", "freundliche", "freundlichen",
  "danke", "mail", "regards", "dear", "hello", "thanks",
]);

export interface MailKeywords { subjectWords: string[]; bodyWords: string[]; }

function tokenize(text: string): string[] {
  return text
    .toLowerCase()
    .split(/[^\p{L}\p{N}]+/u)
    .filter((w) => w.length >= 3 && !STOPWORDS.has(w));
}

export function extractMailKeywords(subject: string, bodyText: string): MailKeywords {
  const cleanSubject = subject.replace(/^((re|aw|fw|fwd|wg)\s*:\s*)+/i, "");
  const subjectWords = [...new Set(tokenize(cleanSubject))];
  const subjectSet = new Set(subjectWords);
  const bodyWords = [...new Set(tokenize(bodyText.slice(0, 2000)))].filter((w) => !subjectSet.has(w));
  return { subjectWords, bodyWords };
}

// Top-3-Vorschlaege: Betreff-Wort im Titel = 3 Punkte, Body-Wort = 1 Punkt.
// Array.prototype.sort ist stabil, Gleichstand behaelt also die Task-Reihenfolge.
export function suggestTasks(tasks: TodoistTask[], kw: MailKeywords): TodoistTask[] {
  const scored = tasks
    .map((task) => {
      const hay = task.content.toLowerCase();
      let score = 0;
      for (const w of kw.subjectWords) if (hay.includes(w)) score += 3;
      for (const w of kw.bodyWords) if (hay.includes(w)) score += 1;
      return { task, score };
    })
    .filter((s) => s.score > 0);
  scored.sort((a, b) => b.score - a.score);
  return scored.slice(0, 3).map((s) => s.task);
}

// Pfeiltasten-Auswahl: clampt auf [0, count-1]; ohne Zeilen gibt es keine Auswahl (-1).
export function moveSelection(current: number, delta: number, count: number): number {
  if (count <= 0) return -1;
  return Math.min(Math.max(current + delta, 0), count - 1);
}

export type DueChip = "today" | "tomorrow" | "nextWeek" | "none";

export interface NewTaskInput {
  title: string;
  projectId: string | null;
  priority: number; // 1..4 wie Todoist, 1 = keine (P4)
  dueChip: DueChip;
  dueText: string;
  today: string; // YYYY-MM-DD
}

function addDays(iso: string, days: number): string {
  const d = new Date(`${iso}T12:00:00Z`);
  d.setUTCDate(d.getUTCDate() + days);
  return d.toISOString().slice(0, 10);
}

// Freitext gewinnt ueber Chip (geht als due_string an Todoists Sprach-Parser);
// Chips erzeugen deterministische due_date ohne Parsing-Risiko.
export function buildNewTaskOptions(input: NewTaskInput): NewTaskOptions {
  const opts: NewTaskOptions = { content: input.title.trim() || "Mail" };
  if (input.projectId) opts.project_id = input.projectId;
  if (input.priority > 1) opts.priority = input.priority;
  const dueText = input.dueText.trim();
  if (dueText) {
    opts.due_string = dueText;
    opts.due_lang = "de";
  } else if (input.dueChip === "today") {
    opts.due_date = input.today;
  } else if (input.dueChip === "tomorrow") {
    opts.due_date = addDays(input.today, 1);
  } else if (input.dueChip === "nextWeek") {
    opts.due_date = addDays(input.today, 7);
  }
  return opts;
}
