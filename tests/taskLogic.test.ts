import { todayIso, groupTasks, priorityColor, taskDeepLink, filterTasks, dueTodayOrOverdue, extractMailKeywords, suggestTasks, moveSelection, buildNewTaskOptions } from "../src/taskpane/taskLogic";
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
  test("baut todoist://-Desktop-URL", () => {
    expect(taskDeepLink("t5")).toBe("todoist://task?id=t5");
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

  test("#-Wort matcht nur den Projektnamen, nicht den Titel", () => {
    const tasks = [named("imTitel", "sap Lizenzen klären", "p1"), named("imProjekt", "Rechnung zahlen", "p2")];
    expect(filterTasks(tasks, "#sap", projects).map((x) => x.id)).toEqual(["imProjekt"]);
  });

  test("#Projekt kombiniert mit normalem Wort (UND)", () => {
    const tasks = [named("a", "Rechnung pruefen", "p2"), named("b", "Schulung planen", "p2"), named("c", "Rechnung zahlen", "p1")];
    expect(filterTasks(tasks, "#sap rechnung", projects).map((x) => x.id)).toEqual(["a"]);
  });

  test("nacktes # wird ignoriert", () => {
    const tasks = [named("a", "x"), named("b", "y")];
    expect(filterTasks(tasks, "#", projects)).toHaveLength(2);
  });

  test("#-Match ist case-insensitiv", () => {
    const tasks = [named("a", "Rechnung", "p2")];
    expect(filterTasks(tasks, "#SAP", projects).map((x) => x.id)).toEqual(["a"]);
  });
});

describe("dueTodayOrOverdue", () => {
  test("behaelt heute + ueberfaellig, wirft ohne Datum und Zukunft raus", () => {
    const tasks = [t("alt", "2026-07-01"), t("heute", "2026-07-06"), t("morgen", "2026-07-07"), t("ohne", null)];
    expect(dueTodayOrOverdue(tasks, "2026-07-06").map((x) => x.id)).toEqual(["alt", "heute"]);
  });
});

describe("extractMailKeywords", () => {
  test("entfernt gestapelte RE:/AW:/FW:/WG:-Praefixe", () => {
    const kw = extractMailKeywords("AW: RE: SAP Lizenzverlängerung", "");
    expect(kw.subjectWords).toEqual(["sap", "lizenzverlängerung"]);
  });

  test("filtert Stoppwoerter und Kurzwoerter", () => {
    const kw = extractMailKeywords("Die Offerte für das neue CRM", "");
    expect(kw.subjectWords).toEqual(["offerte", "neue", "crm"]);
  });

  test("Body-Woerter dedupliziert und ohne Betreff-Duplikate", () => {
    const kw = extractMailKeywords("SAP Update", "Das SAP Update betrifft die Migration. Migration startet bald.");
    expect(kw.subjectWords).toEqual(["sap", "update"]);
    expect(kw.bodyWords).toEqual(["betrifft", "migration", "startet", "bald"]);
  });

  test("Body wird bei 2000 Zeichen gekappt", () => {
    const kw = extractMailKeywords("x", "a".repeat(2000) + " spätwort");
    expect(kw.bodyWords).not.toContain("spätwort");
  });

  test("leerer Input liefert leere Listen", () => {
    expect(extractMailKeywords("", "")).toEqual({ subjectWords: [], bodyWords: [] });
  });
});

describe("moveSelection", () => {
  test("normaler Schritt runter und rauf", () => {
    expect(moveSelection(1, 1, 5)).toBe(2);
    expect(moveSelection(2, -1, 5)).toBe(1);
  });
  test("clampt oben und unten", () => {
    expect(moveSelection(0, -1, 5)).toBe(0);
    expect(moveSelection(4, 1, 5)).toBe(4);
  });
  test("leere Liste liefert -1", () => {
    expect(moveSelection(0, 1, 0)).toBe(-1);
  });
  test("negativer Startindex landet bei 0", () => {
    expect(moveSelection(-1, 1, 3)).toBe(0);
  });
});

describe("buildNewTaskOptions", () => {
  const base = { title: "Task", projectId: null, priority: 1, dueChip: "none" as const, dueText: "", today: "2026-07-06" };

  test("minimal: nur content", () => {
    expect(buildNewTaskOptions(base)).toEqual({ content: "Task" });
  });
  test("Chips mappen deterministisch auf due_date (+0/+1/+7)", () => {
    expect(buildNewTaskOptions({ ...base, dueChip: "today" }).due_date).toBe("2026-07-06");
    expect(buildNewTaskOptions({ ...base, dueChip: "tomorrow" }).due_date).toBe("2026-07-07");
    expect(buildNewTaskOptions({ ...base, dueChip: "nextWeek" }).due_date).toBe("2026-07-13");
  });
  test("Monatswechsel beim Addieren", () => {
    expect(buildNewTaskOptions({ ...base, dueChip: "tomorrow", today: "2026-07-31" }).due_date).toBe("2026-08-01");
  });
  test("Freitext gewinnt ueber Chip und geht als due_string mit de", () => {
    const o = buildNewTaskOptions({ ...base, dueChip: "today", dueText: "naechsten freitag" });
    expect(o.due_string).toBe("naechsten freitag");
    expect(o.due_lang).toBe("de");
    expect(o.due_date).toBeUndefined();
  });
  test("Prioritaet 1 (keine) wird weggelassen, andere gesendet", () => {
    expect(buildNewTaskOptions(base).priority).toBeUndefined();
    expect(buildNewTaskOptions({ ...base, priority: 4 }).priority).toBe(4);
  });
  test("leerer Titel faellt auf Mail zurueck, projectId wird uebernommen", () => {
    const o = buildNewTaskOptions({ ...base, title: "   ", projectId: "p7" });
    expect(o.content).toBe("Mail");
    expect(o.project_id).toBe("p7");
  });
});

describe("suggestTasks", () => {
  const kw = { subjectWords: ["sap", "lizenz"], bodyWords: ["migration"] };

  test("gewichtet Betreff 3, Body 1 und sortiert absteigend", () => {
    const tasks = [
      named("nurBody", "Migration vorbereiten"),
      named("beide", "SAP Lizenz verlängern"),
      named("einSubject", "SAP Schulung"),
    ];
    expect(suggestTasks(tasks, kw).map((x) => x.id)).toEqual(["beide", "einSubject", "nurBody"]);
  });

  test("Score 0 wird nie vorgeschlagen", () => {
    expect(suggestTasks([named("a", "Zmittag buchen")], kw)).toEqual([]);
  });

  test("maximal 3 Vorschlaege", () => {
    const tasks = ["a", "b", "c", "d"].map((id) => named(id, `SAP ${id}`));
    expect(suggestTasks(tasks, kw)).toHaveLength(3);
  });

  test("matcht case-insensitiv gegen den Task-Titel", () => {
    expect(suggestTasks([named("a", "sap lizenzen")], kw).map((x) => x.id)).toEqual(["a"]);
  });

  test("stabile Reihenfolge bei Gleichstand", () => {
    const tasks = [named("erst", "SAP eins"), named("zweit", "SAP zwei")];
    expect(suggestTasks(tasks, kw).map((x) => x.id)).toEqual(["erst", "zweit"]);
  });
});
