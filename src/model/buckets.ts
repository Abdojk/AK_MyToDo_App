import { civilDayNumber } from './dates';
import type { Bucket, HubItem, HubTask } from './types';
import { BUCKET_ORDER } from './types';

/**
 * Place a date on the timeline by comparing calendar days, so a task due at 23:00
 * local time stays in Today rather than sliding into Tomorrow.
 *
 * `floating` marks a value whose midnight belongs to no zone — an all-day event,
 * which Graph stores at midnight and which must not be converted into the
 * viewer's zone, or it moves a day for any zone behind UTC.
 */
export function bucketOf(
  date: Date | null,
  now: Date,
  zone: string,
  floating = false,
): Bucket {
  if (!date) return 'noDueDate';

  const day = civilDayNumber(date, floating ? 'UTC' : zone);
  const today = civilDayNumber(now, zone);

  if (day < today) return 'overdue';
  if (day === today) return 'today';
  if (day === today + 1) return 'tomorrow';
  if (day <= today + 7) return 'next7';
  return 'later';
}

/** A task the Hub no longer shows: finished work. */
export function isClosed(task: HubTask): boolean {
  return task.status === 'completed';
}

/** The date an item sits on, and whether that date floats. */
function anchorOf(item: HubItem): { date: Date | null; floating: boolean } {
  if (item.kind === 'event') {
    return { date: item.event.start, floating: item.event.isAllDay };
  }
  return { date: item.task.due, floating: false };
}

/** Finished work leaves the timeline, whichever source it came from. */
export function isDone(item: HubItem): boolean {
  if (item.kind === 'task') return isClosed(item.task);
  if (item.kind === 'planner') return item.task.percentComplete >= 100;
  return false;
}

const IMPORTANCE_RANK = { high: 0, normal: 1, low: 2 } as const;

function compare(a: HubItem, b: HubItem): number {
  const aTime = anchorOf(a).date?.getTime() ?? Number.POSITIVE_INFINITY;
  const bTime = anchorOf(b).date?.getTime() ?? Number.POSITIVE_INFINITY;
  if (aTime !== bTime) return aTime - bTime;

  // A task's due date sits at local midnight, so on a shared day it already
  // heads the column. Equal timestamps put tasks of either source before an event.
  if (a.kind !== b.kind) {
    if (a.kind === 'event') return 1;
    if (b.kind === 'event') return -1;
    return a.kind === 'task' ? -1 : 1;
  }

  if (a.kind === 'task' && b.kind === 'task') {
    const rank =
      IMPORTANCE_RANK[a.task.importance] - IMPORTANCE_RANK[b.task.importance];
    if (rank !== 0) return rank;
    return a.task.title.localeCompare(b.task.title);
  }

  if (a.kind === 'event' && b.kind === 'event') {
    return a.event.subject.localeCompare(b.event.subject);
  }

  if (a.kind === 'planner' && b.kind === 'planner') {
    return a.task.title.localeCompare(b.task.title);
  }

  return 0;
}

/** Group open tasks and calendar occurrences into the six timeline columns. */
export function groupItems(
  items: HubItem[],
  now: Date,
  zone: string,
): Record<Bucket, HubItem[]> {
  const grouped = Object.fromEntries(
    BUCKET_ORDER.map((b) => [b, [] as HubItem[]]),
  ) as Record<Bucket, HubItem[]>;

  for (const item of items) {
    if (isDone(item)) continue;
    const { date, floating } = anchorOf(item);
    grouped[bucketOf(date, now, zone, floating)].push(item);
  }

  for (const bucket of BUCKET_ORDER) grouped[bucket].sort(compare);

  return grouped;
}
