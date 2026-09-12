import { itemKey } from './types';
import type { HubItem, HubTask } from './types';

/** Replace one item by its key, leaving the rest of the timeline untouched. */
export function applyItem(items: HubItem[], next: HubItem): HubItem[] {
  const key = itemKey(next);
  return items.map((i) => (itemKey(i) === key ? next : i));
}

/**
 * Fold a To Do PATCH response into the task the timeline holds.
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
