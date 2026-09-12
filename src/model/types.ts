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
