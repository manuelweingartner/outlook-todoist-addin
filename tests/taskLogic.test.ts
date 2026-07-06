import { todayIso, groupTasks, priorityColor, taskDeepLink, filterTasks, dueTodayOrOverdue } from "../src/taskpane/taskLogic";
import { TodoistTask } from "../src/lib/todoist";

const t = (id: string, due: string | null, priority = 1): TodoistTask => ({
  id, content: id, project_id: "p1", priority, due: due ? { date: due } : null,
});

const named = (id: string, content: string, project_id = "p1"): TodoistTask => ({
  id, content, project_id, priority: 1, due: null,
});

describe("todayIso", () => {
  test("formatiert YYYY-MM-DD", () => {
    expect(todayIso(new Date("2026-06-29T12:00:00Z"))).toBe("2026-06-29");
  });
});

describe("groupTasks", () => {
  test("trennt ueberfaellig (vor heute) von heute/ohne Datum", () => {
    const g = groupTasks([t("a", "2026-06-28"), t("b", "2026-06-29"), t("c", null)], "2026-06-29");
    expect(g.overdue.map((x) => x.id)).toEqual(["a"]);
    expect(g.today.map((x) => x.id)).toEqual(["b", "c"]);
  });
});

describe("priorityColor", () => {
  test("P1..P4 Mapping", () => {
    expect(priorityColor(4)).toBe("#e44332");
    expect(priorityColor(3)).toBe("#eb8909");
    expect(priorityColor(2)).toBe("#246fe0");
    expect(priorityColor(1)).toBe("#808080");
    expect(priorityColor(undefined)).toBe("#808080");
  });
});

describe("taskDeepLink", () => {
  test("baut Todoist-Task-URL", () => {
    expect(taskDeepLink("t5")).toBe("https://app.todoist.com/app/task/t5");
  });
});

describe("filterTasks", () => {
  const projects = { p1: "Inbox", p2: "SAP" };

  test("matcht case-insensitiv: 'sap' findet 'SAP'", () => {
    const tasks = [named("a", "SAP Lizenzen klären"), named("b", "Zmittag buchen")];
    expect(filterTasks(tasks, "sap", projects).map((x) => x.id)).toEqual(["a"]);
  });

  test("mehrere Woerter sind UND-verknuepft", () => {
    const tasks = [named("a", "SAP Lizenzen klären"), named("b", "SAP Schulung planen")];
    expect(filterTasks(tasks, "sap schulung", projects).map((x) => x.id)).toEqual(["b"]);
  });

  test("matcht auch den Projektnamen", () => {
    const tasks = [named("a", "Rechnung prüfen", "p2"), named("b", "Rechnung zahlen", "p1")];
    expect(filterTasks(tasks, "sap", projects).map((x) => x.id)).toEqual(["a"]);
  });

  test("leere Query liefert alle Tasks", () => {
    const tasks = [named("a", "x"), named("b", "y")];
    expect(filterTasks(tasks, "  ", projects)).toHaveLength(2);
  });

  test("unbekannte project_id crasht nicht", () => {
    const tasks = [named("a", "SAP Thema", "p99")];
    expect(filterTasks(tasks, "sap", projects).map((x) => x.id)).toEqual(["a"]);
  });
});

describe("dueTodayOrOverdue", () => {
  test("behaelt heute + ueberfaellig, wirft ohne Datum und Zukunft raus", () => {
    const tasks = [t("alt", "2026-07-01"), t("heute", "2026-07-06"), t("morgen", "2026-07-07"), t("ohne", null)];
    expect(dueTodayOrOverdue(tasks, "2026-07-06").map((x) => x.id)).toEqual(["alt", "heute"]);
  });
});
