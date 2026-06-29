import { TodoistTask } from "../lib/todoist";

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

export function taskDeepLink(id: string): string {
  return `https://app.todoist.com/app/task/${id}`;
}
