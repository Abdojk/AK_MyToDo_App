import type {
  HubPlannerTask,
  PlannerPriority,
  RawPlannerTask,
} from './types';

/**
 * Planner's own reading of the 0-10 scale, quoted from the plannerTask
 * documentation: values 0 and 1 are "urgent", 2, 3 and 4 "important", 5, 6 and 7
 * "medium", and 8, 9 and 10 "low".
 */
export function plannerPriority(priority: number | null | undefined): PlannerPriority {
  const value = typeof priority === 'number' ? priority : 5;
  if (value <= 1) return 'urgent';
  if (value <= 4) return 'important';
  if (value <= 7) return 'medium';
  return 'low';
}

/** Only the two bands worth interrupting the column for. */
export function isPressing(priority: PlannerPriority): boolean {
  return priority === 'urgent' || priority === 'important';
}

/**
 * A link to the task in Planner.
 *
 * Microsoft Learn documents no deep-link URL for a plannerTask, so this pattern
 * is not verified. Confirm it once against the tenant; correcting it is a change
 * to this one function.
 */
export function plannerTaskUrl(tenantId: string | undefined, taskId: string): string {
  if (!tenantId) return '';
  return `https://tasks.office.com/${encodeURIComponent(tenantId)}/Home/Task/${encodeURIComponent(taskId)}`;
}

/** True once the task is finished, which takes it off the timeline. */
export function isPlannerComplete(raw: RawPlannerTask): boolean {
  return (raw.percentComplete ?? 0) >= 100;
}

/**
 * Map a Graph plannerTask onto the shape the timeline renders.
 *
 * Returns null for a task with no due date. Planner backlogs are largely undated,
 * and the No due date column exists for flagged email, so undated Planner work
 * stays out of the Hub rather than burying it.
 */
export function normalisePlannerTask(
  raw: RawPlannerTask,
  planTitles: Map<string, string>,
  tenantId: string | undefined,
): HubPlannerTask | null {
  if (!raw.dueDateTime) return null;
  const due = new Date(raw.dueDateTime);
  if (Number.isNaN(due.getTime())) return null;
  if (isPlannerComplete(raw)) return null;

  const planId = raw.planId ?? '';
  const total = raw.checklistItemCount ?? 0;
  const active = raw.activeChecklistItemCount ?? 0;

  return {
    id: raw.id,
    title: raw.title?.trim() || '(untitled task)',
    due,
    percentComplete: raw.percentComplete ?? 0,
    priority: plannerPriority(raw.priority),
    planId,
    planTitle: planTitles.get(planId) ?? null,
    checklistDone: Math.max(0, total - active),
    checklistTotal: total,
    etag: raw['@odata.etag'] ?? null,
    url: plannerTaskUrl(tenantId, raw.id),
  };
}

/** A Planner due date goes back as a bare ISO 8601 instant in UTC. */
export function toPlannerDueDate(instant: Date): string {
  return instant.toISOString();
}
