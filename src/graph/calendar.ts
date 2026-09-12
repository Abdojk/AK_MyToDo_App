import { GraphError } from './client';
import type { GraphClient } from './client';
import { normaliseEvents } from '../model/events';
import { presetDate } from '../model/reschedule';
import type { HubEvent, RawEvent } from '../model/types';

/** How far ahead the timeline reads the calendar. */
export const HORIZON_DAYS = 30;

const EVENT_SELECT = [
  'id',
  'subject',
  'start',
  'end',
  'isAllDay',
  'isCancelled',
  'isOnlineMeeting',
  'isOrganizer',
  'location',
  'organizer',
  'responseStatus',
  'showAs',
  'type',
  'webLink',
].join(',');

const MS_PER_DAY = 86_400_000;

/** The window the view covers: local midnight today, to the horizon. */
export function calendarWindow(
  now: Date,
  zone: string,
): { from: Date; to: Date } {
  const from = presetDate('today', now, zone);
  return { from, to: new Date(from.getTime() + HORIZON_DAYS * MS_PER_DAY) };
}

/**
 * Read the default calendar over the window.
 *
 * calendarView expands a recurring series into its occurrences, so the Hub never
 * walks a recurrence pattern. No Prefer header goes out, which means start and
 * end arrive in UTC, the form resolveGraphDateTime handles.
 *
 * Like the sender enrichment, this call is optional: a rejection returns null and
 * the task timeline renders unchanged. A tenant that withholds calendar consent
 * must not cost the Hub its tasks.
 */
export async function fetchCalendarEvents(
  client: GraphClient,
  now: Date,
  zone: string,
): Promise<HubEvent[] | null> {
  const { from, to } = calendarWindow(now, zone);
  const query =
    `/me/calendarView?startDateTime=${encodeURIComponent(from.toISOString())}` +
    `&endDateTime=${encodeURIComponent(to.toISOString())}` +
    `&$select=${EVENT_SELECT}&$orderby=start/dateTime&$top=100`;

  try {
    return normaliseEvents(await client.getAll<RawEvent>(query));
  } catch (e) {
    if (e instanceof GraphError) {
      console.warn(
        `[graph] calendar unavailable (${e.status} ${e.code ?? ''}). ` +
          'The timeline shows tasks only.',
      );
      return null;
    }
    throw e;
  }
}
