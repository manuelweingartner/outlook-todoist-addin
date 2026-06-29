import { todayIso, groupTasks, priorityColor, taskDeepLink } from "../src/taskpane/taskLogic";
import { TodoistTask } from "../src/lib/todoist";

const t = (id: string, due: string | null, priority = 1): TodoistTask => ({
  id, content: id, project_id: "p1", priority, due: due ? { date: due } : null,
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
