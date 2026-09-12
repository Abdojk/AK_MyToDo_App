import { GraphError } from './client';
import type { GraphClient } from './client';
import {
  isPlannerComplete,
  normalisePlannerTask,
  toPlannerDueDate,
} from '../model/planner';
import type {
  HubPlannerTask,
  RawPlannerPlan,
  RawPlannerTask,
} from '../model/types';

/** Every Planner write needs the last known ETag, and asks for the new task back. */
function writeHeaders(etag: string | null): Record<string, string> {
  return {
    'If-Match': etag ?? '*',
    Prefer: 'return=representation',
  };
}

function taskPath(taskId: string): string {
  return `/planner/tasks/${taskId}`;
}

/** Plan titles for the plans the returned tasks belong to, fetched in one batch. */
async function fetchPlanTitles(
  client: GraphClient,
  planIds: string[],
): Promise<Map<string, string>> {
  const titles = new Map<string, string>();
  if (planIds.length === 0) return titles;

  const requests = planIds.map((id, index) => ({
    id: String(index),
    method: 'GET' as const,
    url: `/planner/plans/${id}`,
  }));

  const responses = await client.batchGet<RawPlannerPlan>(requests);
  planIds.forEach((id, index) => {
    const title = responses.get(String(index))?.title?.trim();
    if (title) titles.set(id, title);
  });

  return titles;
}

/**
 * Read the Planner tasks assigned to the signed-in user.
 *
 * Optional in the same way as the calendar and the sender enrichment: a rejection
 * returns null, the App shows a banner, and the rest of the timeline renders.
 */
export async function fetchPlannerTasks(
  client: GraphClient,
  tenantId: string | undefined,
): Promise<HubPlannerTask[] | null> {
  try {
    const raw = await client.getAll<RawPlannerTask>('/me/planner/tasks');
    const open = raw.filter((t) => t.dueDateTime && !isPlannerComplete(t));

    const planIds = [...new Set(open.map((t) => t.planId).filter(Boolean))] as string[];
    const titles = await fetchPlanTitles(client, planIds);

    return open
      .map((t) => normalisePlannerTask(t, titles, tenantId))
      .filter((t): t is HubPlannerTask => t !== null);
  } catch (e) {
    if (e instanceof GraphError) {
      console.warn(
        `[graph] Planner unavailable (${e.status} ${e.code ?? ''}). ` +
          'The timeline shows flagged email and meetings only.',
      );
      return null;
    }
    throw e;
  }
}

/** Re-read one task, for a fresh ETag or after a 204 that carried no body. */
function reread(client: GraphClient, taskId: string): Promise<RawPlannerTask> {
  return client.get<RawPlannerTask>(taskPath(taskId));
}

/**
 * Apply a Planner write.
 *
 * Planner requires If-Match and answers 412 when the ETag is stale, which happens
 * whenever the task changed in Planner since the Hub last read it. One re-read and
 * retry settles that; a second failure reaches the caller so the App rolls back.
 * A 204 means the Prefer header was not honoured, so the task is re-read rather
 * than guessed at.
 */
async function writeTask(
  client: GraphClient,
  task: HubPlannerTask,
  body: Record<string, unknown>,
  tenantId: string | undefined,
  planTitles: Map<string, string>,
): Promise<HubPlannerTask | null> {
  let etag = task.etag;
  let raw: RawPlannerTask | null;

  try {
    raw = await client.patch<RawPlannerTask>(taskPath(task.id), body, writeHeaders(etag));
  } catch (e) {
    if (!(e instanceof GraphError) || e.status !== 412) throw e;
    const fresh = await reread(client, task.id);
    etag = fresh['@odata.etag'] ?? null;
    raw = await client.patch<RawPlannerTask>(taskPath(task.id), body, writeHeaders(etag));
  }

  if (!raw) raw = await reread(client, task.id);
  return normalisePlannerTask(
    { ...raw, planId: raw.planId ?? task.planId },
    planTitles,
    tenantId,
  );
}

function titlesOf(task: HubPlannerTask): Map<string, string> {
  return task.planTitle
    ? new Map([[task.planId, task.planTitle]])
    : new Map<string, string>();
}

/** Mark the task finished. percentComplete 100 is Planner's definition of done. */
export function completePlannerTask(
  client: GraphClient,
  task: HubPlannerTask,
  tenantId: string | undefined,
): Promise<HubPlannerTask | null> {
  return writeTask(client, task, { percentComplete: 100 }, tenantId, titlesOf(task));
}

/** Reopen a task completed by mistake. */
export function reopenPlannerTask(
  client: GraphClient,
  task: HubPlannerTask,
  tenantId: string | undefined,
): Promise<HubPlannerTask | null> {
  return writeTask(client, task, { percentComplete: 0 }, tenantId, titlesOf(task));
}

/** Move the task's due date. Planner takes a bare DateTimeOffset in UTC. */
export function reschedulePlannerTask(
  client: GraphClient,
  task: HubPlannerTask,
  due: Date,
  tenantId: string | undefined,
): Promise<HubPlannerTask | null> {
  return writeTask(
    client,
    task,
    { dueDateTime: toPlannerDueDate(due) },
    tenantId,
    titlesOf(task),
  );
}
