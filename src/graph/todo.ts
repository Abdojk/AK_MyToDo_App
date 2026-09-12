import type { GraphClient } from './client';
import { normaliseTask } from '../model/normalise';
import type { TaskListRef } from '../model/normalise';
import type {
  HubTask,
  RawLinkedResource,
  RawTodoTask,
  RawTodoTaskList,
} from '../model/types';

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

/**
 * Every To Do list in the mailbox.
 *
 * The built-in Flagged email list is one of them, marked by its well-known name.
 * The rest are the built-in Tasks list and whatever lists the user has created —
 * the same set the Planner app in Teams shows under My Tasks and My Plans.
 */
export async function getTaskLists(client: GraphClient): Promise<TaskListRef[]> {
  const lists = await client.getAll<RawTodoTaskList>(
    '/me/todo/lists?$select=id,displayName,wellknownListName',
  );
  return lists.map((list) => ({
    id: list.id,
    displayName: list.displayName?.trim() || 'Tasks',
    isFlaggedEmail: list.wellknownListName === 'flaggedEmails',
  }));
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
  const byTask = new Map<string, RawLinkedResource[]>();
  for (const task of tasks) {
    if (task.linkedResources) byTask.set(task.id, task.linkedResources);
  }

  const pending = tasks.filter(
    (t) => t.linkedResources === undefined || t.linkedResources === null,
  );
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

/** Every task in one list, normalised and carrying that list's identity. */
export async function loadList(
  client: GraphClient,
  list: TaskListRef,
): Promise<HubTask[]> {
  const raw = await getTasks(client, list.id);
  const links = await attachLinkedResources(client, list.id, raw);
  return raw.map((task) => normaliseTask(task, list, links.get(task.id) ?? []));
}

/**
 * The whole To Do read: every list, its tasks, and their links.
 *
 * Each task carries its own listId, because a write has to reach the list the
 * task actually lives in.
 */
export async function loadTasks(client: GraphClient): Promise<HubTask[]> {
  const lists = await getTaskLists(client);
  const perList = await Promise.all(lists.map((list) => loadList(client, list)));
  return perList.flat();
}
