import type { HubTask } from './types';

/** Replace one task by id, leaving the rest of the list untouched. */
export function applyTask(tasks: HubTask[], next: HubTask): HubTask[] {
  return tasks.map((t) => (t.id === next.id ? next : t));
}

/**
 * Fold a PATCH response into the task the timeline holds.
 *
 * The response carries the task's own fields but no linked resources and no
 * message data, so the fields that only the read supplies are carried across
 * rather than blanked.
 */
export function mergeWritten(previous: HubTask, fresh: HubTask): HubTask {
  return {
    ...fresh,
    outlookUrl: fresh.outlookUrl ?? previous.outlookUrl,
    sender: previous.sender,
    receivedDateTime: previous.receivedDateTime,
  };
}
