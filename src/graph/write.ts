import { GraphError } from './client';
import type { GraphClient } from './client';
import { toGraphDate } from '../model/reschedule';
import type { RawTodoTask } from '../model/types';

function taskPath(listId: string, taskId: string): string {
  return `/me/todo/lists/${listId}/tasks/${taskId}`;
}

/**
 * Update todoTask is documented to answer 200 with the updated task, so an empty
 * body means something changed at the service and the caller should not guess.
 */
function required(raw: RawTodoTask | null): RawTodoTask {
  if (!raw) {
    throw new Error(
      'Microsoft Graph accepted the change but returned no task. Refresh to see the current state.',
    );
  }
  return raw;
}

/**
 * Mark a task finished.
 *
 * Whether Graph accepts `status` on its own, or requires `completedDateTime`
 * alongside it, is not stated in the Update todoTask documentation. The status-only
 * body goes first; a 400 triggers one retry carrying both.
 */
export async function completeTask(
  client: GraphClient,
  listId: string,
  taskId: string,
): Promise<RawTodoTask> {
  const path = taskPath(listId, taskId);
  try {
    return required(await client.patch<RawTodoTask>(path, { status: 'completed' }));
  } catch (e) {
    if (e instanceof GraphError && e.status === 400) {
      return required(
        await client.patch<RawTodoTask>(path, {
          status: 'completed',
          completedDateTime: toGraphDate(new Date()),
        }),
      );
    }
    throw e;
  }
}

/** Reopen a task completed by mistake. */
export async function reopenTask(
  client: GraphClient,
  listId: string,
  taskId: string,
): Promise<RawTodoTask> {
  return required(
    await client.patch<RawTodoTask>(taskPath(listId, taskId), {
      status: 'notStarted',
    }),
  );
}

/** Move a task's due date. Graph accepts dueDateTime on its own. */
export async function rescheduleTask(
  client: GraphClient,
  listId: string,
  taskId: string,
  due: Date,
): Promise<RawTodoTask> {
  return required(
    await client.patch<RawTodoTask>(taskPath(listId, taskId), {
      dueDateTime: toGraphDate(due),
    }),
  );
}
