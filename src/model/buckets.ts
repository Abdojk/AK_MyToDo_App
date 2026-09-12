import { civilDayNumber } from './dates';
import type { Bucket, HubTask } from './types';
import { BUCKET_ORDER } from './types';

/**
 * Place a due date on the timeline by comparing calendar days in `zone`, so a
 * task due at 23:00 local time stays in Today rather than sliding into Tomorrow.
 */
export function bucketOf(
  due: Date | null,
  now: Date,
  zone: string,
): Bucket {
  if (!due) return 'noDueDate';

  const dueDay = civilDayNumber(due, zone);
  const today = civilDayNumber(now, zone);

  if (dueDay < today) return 'overdue';
  if (dueDay === today) return 'today';
  if (dueDay === today + 1) return 'tomorrow';
  if (dueDay <= today + 7) return 'next7';
  return 'later';
}

/** A task the Hub no longer shows: finished work. */
export function isClosed(task: HubTask): boolean {
  return task.status === 'completed';
}

/** Group open tasks into the six timeline columns, each sorted by due date. */
export function groupByBucket(
  tasks: HubTask[],
  now: Date,
  zone: string,
): Record<Bucket, HubTask[]> {
  const grouped = Object.fromEntries(
    BUCKET_ORDER.map((b) => [b, [] as HubTask[]]),
  ) as Record<Bucket, HubTask[]>;

  for (const task of tasks) {
    if (isClosed(task)) continue;
    grouped[bucketOf(task.due, now, zone)].push(task);
  }

  const importanceRank = { high: 0, normal: 1, low: 2 } as const;
  for (const bucket of BUCKET_ORDER) {
    grouped[bucket].sort((a, b) => {
      const aTime = a.due?.getTime() ?? Number.POSITIVE_INFINITY;
      const bTime = b.due?.getTime() ?? Number.POSITIVE_INFINITY;
      if (aTime !== bTime) return aTime - bTime;
      const rank = importanceRank[a.importance] - importanceRank[b.importance];
      if (rank !== 0) return rank;
      return a.title.localeCompare(b.title);
    });
  }

  return grouped;
}
