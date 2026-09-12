/** Timeline column a task falls into. */
export type Bucket =
  | 'overdue'
  | 'today'
  | 'tomorrow'
  | 'next7'
  | 'later'
  | 'noDueDate';

export const BUCKET_ORDER: Bucket[] = [
  'overdue',
  'today',
  'tomorrow',
  'next7',
  'later',
  'noDueDate',
];

export const BUCKET_LABEL: Record<Bucket, string> = {
  overdue: 'Overdue',
  today: 'Today',
  tomorrow: 'Tomorrow',
  next7: 'Next 7 days',
  later: 'Later',
  noDueDate: 'No due date',
};

/** Microsoft Graph dateTimeTimeZone: a wall-clock string plus the zone it belongs to. */
export interface GraphDateTimeTimeZone {
  dateTime: string;
  timeZone: string;
}

/** Microsoft Graph linkedResource, trimmed to the fields the Hub reads. */
export interface RawLinkedResource {
  id: string;
  applicationName?: string | null;
  displayName?: string | null;
  externalId?: string | null;
  webUrl?: string | null;
}

/** Microsoft Graph todoTask, trimmed to the fields the Hub reads. */
export interface RawTodoTask {
  id: string;
  title?: string | null;
  body?: { content?: string | null; contentType?: string | null } | null;
  status?: string | null;
  importance?: string | null;
  categories?: string[] | null;
  dueDateTime?: GraphDateTimeTimeZone | null;
  reminderDateTime?: GraphDateTimeTimeZone | null;
  completedDateTime?: GraphDateTimeTimeZone | null;
  createdDateTime?: string | null;
  lastModifiedDateTime?: string | null;
  linkedResources?: RawLinkedResource[] | null;
}

/** Microsoft Graph todoTaskList, trimmed to the fields the Hub reads. */
export interface RawTodoTaskList {
  id: string;
  displayName?: string | null;
  wellknownListName?: string | null;
}

export type Importance = 'low' | 'normal' | 'high';

/** A task as the Hub renders it. */
export interface HubTask {
  id: string;
  title: string;
  /** The To Do list it lives in. Every write needs it. */
  listId: string;
  listName: string;
  /** True for the built-in Flagged email list, which drives the From line. */
  isFlaggedEmail: boolean;
  due: Date | null;
  status: string;
  importance: Importance;
  categories: string[];
  bodyPreview: string;
  outlookUrl: string | null;
  /** Enrichment only. Null when the mail read is off or unavailable. */
  sender: string | null;
  /** Enrichment only. Null when the mail read is off or unavailable. */
  receivedDateTime: Date | null;
  lastModified: Date | null;
}

/** Microsoft Graph event, trimmed to the fields the Hub reads. */
export interface RawEvent {
  id: string;
  subject?: string | null;
  start?: GraphDateTimeTimeZone | null;
  end?: GraphDateTimeTimeZone | null;
  isAllDay?: boolean | null;
  isCancelled?: boolean | null;
  isOnlineMeeting?: boolean | null;
  isOrganizer?: boolean | null;
  location?: { displayName?: string | null } | null;
  organizer?: {
    emailAddress?: { name?: string | null; address?: string | null };
  } | null;
  responseStatus?: { response?: string | null } | null;
  showAs?: string | null;
  type?: string | null;
  webLink?: string | null;
}

/** A calendar occurrence as the Hub renders it. */
export interface HubEvent {
  id: string;
  subject: string;
  start: Date;
  end: Date;
  /** An all-day event carries a floating midnight, not an instant in any zone. */
  isAllDay: boolean;
  location: string | null;
  organiser: string | null;
  /** free, tentative, busy, oof, workingElsewhere, unknown. */
  showAs: string;
  isOnlineMeeting: boolean;
  isOrganiser: boolean;
  /** The occurrence belongs to a recurring series. */
  isRecurring: boolean;
  outlookUrl: string;
}

/**
 * One entry on the timeline.
 *
 * The union wraps rather than widens, so HubTask keeps the shape every existing
 * reader and test expects.
 */
export type HubItem =
  | { kind: 'task'; task: HubTask }
  | { kind: 'event'; event: HubEvent }
  | { kind: 'planner'; task: HubPlannerTask };

export function taskItem(task: HubTask): HubItem {
  return { kind: 'task', task };
}

export function eventItem(event: HubEvent): HubItem {
  return { kind: 'event', event };
}

export function plannerItem(task: HubPlannerTask): HubItem {
  return { kind: 'planner', task };
}

/** Stable identity across sources, since ids are only unique within one. */
export function itemKey(item: HubItem): string {
  return item.kind === 'event'
    ? `event:${item.event.id}`
    : `${item.kind}:${item.task.id}`;
}

/** Microsoft Graph plannerTask, trimmed to the fields the Hub reads. */
export interface RawPlannerTask {
  id: string;
  title?: string | null;
  planId?: string | null;
  bucketId?: string | null;
  /** A plain DateTimeOffset in UTC, not the dateTimeTimeZone pair todoTask uses. */
  dueDateTime?: string | null;
  startDateTime?: string | null;
  percentComplete?: number | null;
  priority?: number | null;
  checklistItemCount?: number | null;
  activeChecklistItemCount?: number | null;
  hasDescription?: boolean | null;
  /** Required on every write, as the If-Match header. */
  '@odata.etag'?: string | null;
}

export interface RawPlannerPlan {
  id: string;
  title?: string | null;
}

/** Planner's own reading of the 0-10 priority scale. */
export type PlannerPriority = 'urgent' | 'important' | 'medium' | 'low';

/** A Planner task as the Hub renders it. */
export interface HubPlannerTask {
  id: string;
  title: string;
  /** Never null: undated Planner tasks are filtered out before this point. */
  due: Date;
  percentComplete: number;
  priority: PlannerPriority;
  planId: string;
  planTitle: string | null;
  checklistDone: number;
  checklistTotal: number;
  /** Kept current from each PATCH response, because writes need If-Match. */
  etag: string | null;
  url: string;
}
