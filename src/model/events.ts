import { civilParts, resolveGraphDateTime } from './dates';
import type { HubEvent, RawEvent } from './types';

/**
 * A link that opens the occurrence in the current Outlook web client.
 *
 * The event's own webLink is documented as opening "only earlier versions of
 * Outlook on the web", with the current form given as
 * https://outlook.office365.com/calendar/item/{event-id}, the id URL-encoded.
 * The id is built from here, so the link survives a Calendars.ReadBasic response
 * that omits webLink; webLink stands in only when there is no id.
 */
export function eventOutlookUrl(
  id: string | null | undefined,
  webLink: string | null | undefined,
): string {
  if (id) {
    return `https://outlook.office365.com/calendar/item/${encodeURIComponent(id)}`;
  }
  return webLink ?? '';
}

/**
 * Whether the occurrence belongs on the timeline.
 *
 * Cancelled occurrences and invitations you declined are dropped. Everything else
 * stays, including all-day events, events marked free, and events that have
 * already finished today. Only the exact literal `declined` filters, so an
 * unrecognised responseStatus can never hide an event.
 */
export function isVisibleEvent(raw: RawEvent): boolean {
  if (raw.isCancelled === true) return false;
  if (raw.responseStatus?.response === 'declined') return false;
  return true;
}

function organiserOf(raw: RawEvent): string | null {
  const address = raw.organizer?.emailAddress;
  const name = address?.name?.trim() || address?.address?.trim();
  return name || null;
}

/** Map a Graph event onto the shape the timeline renders. */
export function normaliseEvent(raw: RawEvent): HubEvent | null {
  const start = resolveGraphDateTime(raw.start);
  if (!start) return null;
  const end = resolveGraphDateTime(raw.end) ?? start;

  return {
    id: raw.id,
    subject: raw.subject?.trim() || '(no subject)',
    start,
    end,
    isAllDay: raw.isAllDay === true,
    location: raw.location?.displayName?.trim() || null,
    organiser: organiserOf(raw),
    showAs: raw.showAs ?? 'busy',
    isOnlineMeeting: raw.isOnlineMeeting === true,
    isOrganiser: raw.isOrganizer === true,
    isRecurring: raw.type === 'occurrence' || raw.type === 'exception',
    outlookUrl: eventOutlookUrl(raw.id, raw.webLink),
  };
}

/** Read the whole page of events into the timeline's shape. */
export function normaliseEvents(raw: RawEvent[]): HubEvent[] {
  return raw
    .filter(isVisibleEvent)
    .map(normaliseEvent)
    .filter((e): e is HubEvent => e !== null);
}

const pad = (n: number) => String(n).padStart(2, '0');

/**
 * The clock time to show, for example "09:00".
 *
 * An all-day event is read in UTC, because its midnight is floating: converting
 * it into the viewer's zone would move it.
 */
export function formatEventTime(instant: Date, zone: string, floating: boolean): string {
  const p = civilParts(instant, floating ? 'UTC' : zone);
  return `${pad(p.hour)}:${pad(p.minute)}`;
}

/** "09:00 – 10:00", or "All day". */
export function formatEventRange(event: HubEvent, zone: string): string {
  if (event.isAllDay) return 'All day';
  const from = formatEventTime(event.start, zone, false);
  const to = formatEventTime(event.end, zone, false);
  return from === to ? from : `${from} – ${to}`;
}
