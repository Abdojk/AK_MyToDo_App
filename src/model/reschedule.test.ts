import { describe, expect, it } from 'vitest';
import { fromDateInput, presetDate, toDateInput, toGraphDate } from './reschedule';
import { civilDayNumber, resolveGraphDateTime } from './dates';

// 12 Sep 2026, 22:00 in Asia/Amman (UTC+3) is still 19:00 UTC the same day.
const AMMAN = 'Asia/Amman';
const LATE_EVENING = new Date('2026-09-12T19:00:00Z');

// 12 Sep 2026, 21:00 in Los Angeles (UTC-7) is 13 Sep 04:00 UTC: the local date
// and the UTC date differ, which is the case a naive implementation gets wrong.
const LA = 'America/Los_Angeles';
const LATE_NIGHT = new Date('2026-09-13T04:00:00Z');

describe('presetDate', () => {
  it('resolves today to the local calendar day, not the UTC one', () => {
    const due = presetDate('today', LATE_NIGHT, LA);
    expect(toDateInput(due, LA)).toBe('2026-09-12');
  });

  it('resolves tomorrow and next week by calendar days', () => {
    expect(toDateInput(presetDate('tomorrow', LATE_EVENING, AMMAN), AMMAN)).toBe(
      '2026-09-13',
    );
    expect(toDateInput(presetDate('nextWeek', LATE_EVENING, AMMAN), AMMAN)).toBe(
      '2026-09-19',
    );
  });

  it('rolls over the end of a month', () => {
    const endOfMonth = new Date('2026-09-28T09:00:00Z');
    expect(toDateInput(presetDate('nextWeek', endOfMonth, AMMAN), AMMAN)).toBe(
      '2026-10-05',
    );
  });

  it('lands on the intended day when read back through the bucket maths', () => {
    for (const zone of [AMMAN, LA, 'UTC', 'Asia/Tokyo']) {
      const due = presetDate('tomorrow', LATE_EVENING, zone);
      const roundTripped = resolveGraphDateTime(toGraphDate(due));
      expect(roundTripped).not.toBeNull();
      expect(civilDayNumber(roundTripped!, zone)).toBe(
        civilDayNumber(due, zone),
      );
    }
  });
});

describe('fromDateInput', () => {
  it('reads a typed date as local midnight of that day', () => {
    const due = fromDateInput('2026-10-31', AMMAN);
    expect(due).not.toBeNull();
    expect(toDateInput(due!, AMMAN)).toBe('2026-10-31');
  });

  it('rejects anything that is not an ISO date', () => {
    expect(fromDateInput('', AMMAN)).toBeNull();
    expect(fromDateInput('31/10/2026', AMMAN)).toBeNull();
  });
});

describe('toGraphDate', () => {
  it('states the instant in UTC, with no offset suffix', () => {
    const value = toGraphDate(new Date('2026-09-19T21:00:00Z'));
    expect(value).toEqual({ dateTime: '2026-09-19T21:00:00', timeZone: 'UTC' });
  });

  it('pads every component to two digits', () => {
    const value = toGraphDate(new Date('2026-01-02T03:04:05Z'));
    expect(value.dateTime).toBe('2026-01-02T03:04:05');
  });
});
