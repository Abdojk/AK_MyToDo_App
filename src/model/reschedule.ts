import { civilParts, wallClockToInstant } from './dates';
import type { GraphDateTimeTimeZone } from './types';

export type Preset = 'today' | 'tomorrow' | 'nextWeek';

export const PRESETS: { key: Preset; label: string; days: number }[] = [
  { key: 'today', label: 'Today', days: 0 },
  { key: 'tomorrow', label: 'Tomorrow', days: 1 },
  { key: 'nextWeek', label: 'Next week', days: 7 },
];

const ISO_DATE = /^(\d{4})-(\d{2})-(\d{2})$/;

/**
 * A calendar date, as midnight of that date in `zone`.
 *
 * Anchoring on local midnight means the value reads back as the same calendar day
 * the viewer chose, whatever the offset between their zone and UTC.
 */
function localMidnight(
  year: number,
  month: number,
  day: number,
  zone: string,
): Date {
  return wallClockToInstant(
    { year, month, day, hour: 0, minute: 0, second: 0 },
    zone,
  );
}

/** Today, tomorrow, or a week out, counted in calendar days in `zone`. */
export function presetDate(preset: Preset, now: Date, zone: string): Date {
  const offset = PRESETS.find((p) => p.key === preset)?.days ?? 0;
  const today = civilParts(now, zone);
  // Date.UTC normalises a day number past the end of the month.
  const shifted = new Date(
    Date.UTC(today.year, today.month - 1, today.day + offset),
  );
  return localMidnight(
    shifted.getUTCFullYear(),
    shifted.getUTCMonth() + 1,
    shifted.getUTCDate(),
    zone,
  );
}

/** A value from an `<input type="date">`, as midnight of that date in `zone`. */
export function fromDateInput(value: string, zone: string): Date | null {
  const m = ISO_DATE.exec(value.trim());
  if (!m) return null;
  return localMidnight(Number(m[1]), Number(m[2]), Number(m[3]), zone);
}

/** The value an `<input type="date">` expects, for example "2026-09-20". */
export function toDateInput(instant: Date, zone: string): string {
  const p = civilParts(instant, zone);
  const pad = (n: number) => String(n).padStart(2, '0');
  return `${p.year}-${pad(p.month)}-${pad(p.day)}`;
}

/**
 * The dueDateTime a PATCH body carries.
 *
 * Always stated in UTC. Graph accepts a zone name in the request, but the set of
 * names it accepts is not documented for arbitrary IANA zones, and a UTC instant
 * round-trips through resolveGraphDateTime without depending on that.
 */
export function toGraphDate(instant: Date): GraphDateTimeTimeZone {
  const pad = (n: number) => String(n).padStart(2, '0');
  const dateTime =
    `${instant.getUTCFullYear()}-${pad(instant.getUTCMonth() + 1)}-` +
    `${pad(instant.getUTCDate())}T${pad(instant.getUTCHours())}:` +
    `${pad(instant.getUTCMinutes())}:${pad(instant.getUTCSeconds())}`;
  return { dateTime, timeZone: 'UTC' };
}
