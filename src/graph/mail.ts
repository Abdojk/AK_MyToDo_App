import { GraphError } from './client';
import type { GraphClient } from './client';
import { subjectKey } from '../model/normalise';
import type { HubTask } from '../model/types';

interface RawMessage {
  id: string;
  subject?: string | null;
  receivedDateTime?: string | null;
  webLink?: string | null;
  from?: { emailAddress?: { name?: string | null; address?: string | null } } | null;
}

export interface Enrichment {
  sender: string;
  receivedDateTime: Date | null;
  webLink: string | null;
}

/**
 * Optional layer: sender and received date for each flagged task.
 *
 * Microsoft To Do exposes a flagged email as a task title and body only, with no
 * sender field, so the sender comes from the message itself. Two facts are not
 * confirmed in the Microsoft Learn pages consulted: that /me/messages accepts a
 * filter on flag/flagStatus, and that linkedResource.externalId holds a Graph
 * message id. Both are therefore attempted and neither is depended upon. A
 * rejection disables enrichment and the timeline renders without a sender.
 */
export async function fetchFlaggedMessages(
  client: GraphClient,
): Promise<RawMessage[] | null> {
  const query =
    "/me/messages?$select=id,subject,receivedDateTime,webLink,from" +
    "&$filter=flag/flagStatus eq 'flagged'&$top=100";
  try {
    return await client.getAll<RawMessage>(query);
  } catch (e) {
    if (e instanceof GraphError) {
      console.warn(
        `[graph] sender enrichment unavailable (${e.status} ${e.code ?? ''}). ` +
          'The timeline renders without sender names.',
      );
      return null;
    }
    throw e;
  }
}

function senderOf(message: RawMessage): string {
  const address = message.from?.emailAddress;
  return address?.name?.trim() || address?.address?.trim() || 'Unknown sender';
}

/**
 * Join messages onto tasks by subject.
 *
 * A flagged email becomes a task whose title is the message subject, so the
 * subject is the join key. Subjects that appear on more than one flagged message
 * are ambiguous and are skipped rather than guessed at.
 *
 * Only tasks from the Flagged email list are joined. A private task called
 * "Invoice" must not pick up the sender of an unrelated flagged email with the
 * same subject.
 */
export function enrichTasks(tasks: HubTask[], messages: RawMessage[]): HubTask[] {
  const bySubject = new Map<string, RawMessage | 'ambiguous'>();
  for (const message of messages) {
    const key = subjectKey(message.subject ?? '');
    if (!key) continue;
    bySubject.set(key, bySubject.has(key) ? 'ambiguous' : message);
  }

  return tasks.map((task) => {
    if (!task.isFlaggedEmail) return task;
    const match = bySubject.get(subjectKey(task.title));
    if (!match || match === 'ambiguous') return task;
    return {
      ...task,
      sender: senderOf(match),
      receivedDateTime: match.receivedDateTime
        ? new Date(match.receivedDateTime)
        : null,
      outlookUrl: task.outlookUrl ?? match.webLink ?? null,
    };
  });
}
