import type { GraphClient } from './client';
import { normaliseTask } from '../model/normalise';
import type {
  HubTask,
  RawLinkedResource,
  RawTodoTask,
  RawTodoTaskList,
} from '../model/types';

export class FlaggedListMissingError extends Error {
  constructor() {
    super(
      'This mailbox has no built-in Flagged email list in Microsoft To Do. Flag an email in Outlook, then refresh.',
    );
    this.name = 'FlaggedListMissingError';
  }
}

const TASK_SELECT = [
  'id',
  'title',
  'body',
  'status',
  'importance',
  'categories',
  'dueDateTime',
  'reminderDateTime',
  'completedDateTime',
  'createdDateTime',
  'lastModifiedDateTime',
].join(',');

/** Locate the built-in Flagged email list by its well-known name. */
export async function getFlaggedList(
  client: GraphClient,
): Promise<RawTodoTaskList> {
  const lists = await client.getAll<RawTodoTaskList>(
    '/me/todo/lists?$select=id,displayName,wellknownListName',
  );
  const flagged = lists.find((l) => l.wellknownListName === 'flaggedEmails');
  if (!flagged) throw new FlaggedListMissingError();
  return flagged;
}

/**
 * Read every task in a list.
 *
 * $expand=linkedResources is attempted first, because one round trip beats two.
 * The To Do documentation states only that the endpoint supports "some" OData
 * parameters, so a rejection falls back to a batched per-task read.
 */
export async function getTasks(
  client: GraphClient,
  listId: string,
): Promise<RawTodoTask[]> {
  const base = `/me/todo/lists/${listId}/tasks?$top=100&$select=${TASK_SELECT}`;
  try {
    return await client.getAll<RawTodoTask>(`${base}&$expand=linkedResources`);
  } catch (e) {
    console.warn(
      '[graph] $expand=linkedResources rejected, falling back to a batched read.',
      e,
    );
    return client.getAll<RawTodoTask>(base);
  }
}

/** Fetch linked resources for tasks that arrived without them. */
export async function attachLinkedResources(
  client: GraphClient,
  listId: string,
  tasks: RawTodoTask[],
): Promise<Map<string, RawLinkedResource[]>> {
  const pending = tasks.filter((t) => t.linkedResources === undefined || t.linkedResources === null);
  const byTask = new Map<string, RawLinkedResource[]>();
  for (const task of tasks) {
    if (task.linkedResources) byTask.set(task.id, task.linkedResources);
  }
  if (pending.length === 0) return byTask;

  const requests = pending.map((task, index) => ({
    id: String(index),
    method: 'GET' as const,
    url: `/me/todo/lists/${listId}/tasks/${task.id}/linkedResources`,
  }));

  const responses = await client.batchGet<{ value?: RawLinkedResource[] }>(requests);
  pending.forEach((task, index) => {
    byTask.set(task.id, responses.get(String(index))?.value ?? []);
  });

  return byTask;
}

/** The whole read: flagged list, its tasks, their links, normalised for the UI. */
export async function loadFlaggedTasks(client: GraphClient): Promise<HubTask[]> {
  const list = await getFlaggedList(client);
  const raw = await getTasks(client, list.id);
  const links = await attachLinkedResources(client, list.id, raw);
  return raw.map((task) => normaliseTask(task, links.get(task.id) ?? []));
}
