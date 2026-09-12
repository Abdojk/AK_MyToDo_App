import { resolveGraphDateTime } from './dates';
import type { HubTask, Importance, RawLinkedResource, RawTodoTask } from './types';

/** The list a task was read from, carried onto the task for writes and display. */
export interface TaskListRef {
  id: string;
  displayName: string;
  isFlaggedEmail: boolean;
}

function stripHtml(html: string): string {
  return html
    .replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<script[\s\S]*?<\/script>/gi, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\s+/g, ' ')
    .trim();
}

function toPreview(body: RawTodoTask['body'], limit = 220): string {
  const content = body?.content?.trim();
  if (!content) return '';
  const text =
    body?.contentType?.toLowerCase() === 'html'
      ? stripHtml(content)
      : content.replace(/\s+/g, ' ').trim();
  return text.length > limit ? `${text.slice(0, limit - 1)}…` : text;
}

function toImportance(value: string | null | undefined): Importance {
  return value === 'high' || value === 'low' ? value : 'normal';
}

function toDate(value: string | null | undefined): Date | null {
  if (!value) return null;
  const parsed = new Date(value);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

/** Prefer a resource that links back to Outlook, then anything with a URL. */
function pickOutlookUrl(resources: RawLinkedResource[]): string | null {
  const withUrl = resources.filter((r) => Boolean(r.webUrl));
  if (withUrl.length === 0) return null;
  const outlook = withUrl.find((r) =>
    (r.applicationName ?? '').toLowerCase().includes('outlook'),
  );
  return (outlook ?? withUrl[0]).webUrl ?? null;
}

/**
 * Map a Graph todoTask, plus any linked resources fetched for it, onto the shape
 * the timeline renders. Everything the Hub shows is derived here, so the
 * components stay free of Graph field names.
 */
export function normaliseTask(
  raw: RawTodoTask,
  list: TaskListRef,
  linked: RawLinkedResource[] = raw.linkedResources ?? [],
): HubTask {
  return {
    id: raw.id,
    title: raw.title?.trim() || '(no subject)',
    listId: list.id,
    listName: list.displayName,
    isFlaggedEmail: list.isFlaggedEmail,
    due: resolveGraphDateTime(raw.dueDateTime),
    status: raw.status ?? 'notStarted',
    importance: toImportance(raw.importance),
    categories: raw.categories ?? [],
    bodyPreview: toPreview(raw.body),
    outlookUrl: pickOutlookUrl(linked),
    sender: null,
    receivedDateTime: null,
    lastModified: toDate(raw.lastModifiedDateTime),
  };
}

/** Subject text reduced to a comparable key, reply and forward prefixes removed. */
export function subjectKey(subject: string): string {
  return subject
    .replace(/^((re|fw|fwd|aw|tr)\s*(\[\d+\])?\s*:\s*)+/i, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}

/** The list a task belongs to, recovered from the task itself. */
export function listRefOf(task: HubTask): TaskListRef {
  return {
    id: task.listId,
    displayName: task.listName,
    isFlaggedEmail: task.isFlaggedEmail,
  };
}
