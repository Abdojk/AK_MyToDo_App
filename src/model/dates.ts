import type { GraphDateTimeTimeZone } from './types';

const MS_PER_DAY = 86_400_000;

/**
 * Microsoft Graph returns a dueDateTime as a wall-clock string plus a time zone
 * name. The helpers below turn that pair into an absolute instant, and turn an
 * instant into a calendar day number in a chosen zone, so that bucketing
 * compares days rather than raw milliseconds.
 */

interface WallClock {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
}

const GRAPH_DATE_TIME =
  /^(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2}):(\d{2})/;

function parseWallClock(dateTime: string): WallClock | null {
  const m = GRAPH_DATE_TIME.exec(dateTime.trim());
  if (!m) return null;
  return {
    year: Number(m[1]),
    month: Number(m[2]),
    day: Number(m[3]),
    hour: Number(m[4]),
    minute: Number(m[5]),
    second: Number(m[6]),
  };
}

function partsOf(instant: Date, zone: string): WallClock {
  const parts = new Intl.DateTimeFormat('en-US', {
    timeZone: zone,
    hourCycle: 'h23',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    second: '2-digit',
  }).formatToParts(instant);

  const read = (type: Intl.DateTimeFormatPartTypes): number => {
    const part = parts.find((p) => p.type === type);
    return part ? Number(part.value) : 0;
  };

  return {
    year: read('year'),
    month: read('month'),
    day: read('day'),
    hour: read('hour'),
    minute: read('minute'),
    second: read('second'),
  };
}

/** Milliseconds by which the zone leads UTC at the given instant. */
function zoneOffsetMs(instant: Date, zone: string): number {
  const p = partsOf(instant, zone);
  const asIfUtc = Date.UTC(p.year, p.month - 1, p.day, p.hour, p.minute, p.second);
  return asIfUtc - instant.getTime();
}

/** True when Intl accepts the zone name. Graph can return Windows zone names. */
export function isSupportedZone(zone: string): boolean {
  try {
    new Intl.DateTimeFormat('en-US', { timeZone: zone });
    return true;
  } catch {
    return false;
  }
}

/**
 * Resolve a Graph dateTimeTimeZone to an absolute instant.
 *
 * An unrecognised zone name falls back to UTC. Graph returns Windows zone names
 * such as "Pacific Standard Time" in some Outlook responses, and Intl rejects
 * those; UTC is the value Outlook uses for a flag due date by default, so the
 * fallback keeps the day correct in the common case instead of dropping the task.
 */
export function resolveGraphDateTime(
  value: GraphDateTimeTimeZone | null | undefined,
): Date | null {
  if (!value?.dateTime) return null;
  const wall = parseWallClock(value.dateTime);
  if (!wall) return null;

  const asUtc = Date.UTC(
    wall.year,
    wall.month - 1,
    wall.day,
    wall.hour,
    wall.minute,
    wall.second,
  );

  const zone = value.timeZone?.trim();
  if (!zone || zone === 'UTC' || zone === 'Etc/UTC' || zone === 'GMT') {
    return new Date(asUtc);
  }
  if (!isSupportedZone(zone)) return new Date(asUtc);

  // Two passes settle the offset across a daylight-saving boundary.
  let ts = asUtc - zoneOffsetMs(new Date(asUtc), zone);
  ts = asUtc - zoneOffsetMs(new Date(ts), zone);
  return new Date(ts);
}

/** Days since the Unix epoch for the calendar date the instant falls on in `zone`. */
export function civilDayNumber(instant: Date, zone: string): number {
  const p = isSupportedZone(zone)
    ? partsOf(instant, zone)
    : partsOf(instant, 'UTC');
  return Math.floor(Date.UTC(p.year, p.month - 1, p.day) / MS_PER_DAY);
}

/** The viewer's IANA zone, for example "Asia/Amman". */
export function viewerZone(): string {
  return Intl.DateTimeFormat().resolvedOptions().timeZone || 'UTC';
}

const MONTHS = [
  'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec',
];

/**
 * Unambiguous date, for example "12 Sep 2026".
 *
 * Built from parts rather than through a locale, because en-GB renders September
 * as "Sept" and the house format is a three-letter month.
 */
export function formatDueDate(due: Date, zone: string): string {
  const p = isSupportedZone(zone) ? partsOf(due, zone) : partsOf(due, 'UTC');
  return `${p.day} ${MONTHS[p.month - 1]} ${p.year}`;
}
