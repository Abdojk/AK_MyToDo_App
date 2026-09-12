import { describe, expect, it } from 'vitest';
import {
  eventOutlookUrl,
  formatEventRange,
  isVisibleEvent,
  normaliseEvent,
  normaliseEvents,
} from './events';
import { bucketOf, groupItems } from './buckets';
import { eventItem, taskItem } from './types';
import type { HubTask, RawEvent } from './types';

const AMMAN = 'Asia/Amman';       // UTC+3 in September
const LA = 'America/Los_Angeles'; // UTC-7 in September
const NOW = new Date('2026-09-12T07:00:00Z');

const rawEvent = (over: Partial<RawEvent> = {}): RawEvent => ({
  id: 'AAMkAGI2=',
  subject: 'Month-end close review',
  start: { dateTime: '2026-09-12T09:00:00.0000000', timeZone: 'UTC' },
  end: { dateTime: '2026-09-12T10:00:00.0000000', timeZone: 'UTC' },
  isAllDay: false,
  isCancelled: false,
  showAs: 'busy',
  type: 'singleInstance',
  organizer: { emailAddress: { name: 'Rana Haddad', address: 'rana@example.com' } },
  location: { displayName: 'Meeting room 2' },
  ...over,
});

describe('isVisibleEvent', () => {
  it('drops a cancelled occurrence', () => {
    expect(isVisibleEvent(rawEvent({ isCancelled: true }))).toBe(false);
  });

  it('drops an invitation you declined', () => {
    expect(
      isVisibleEvent(rawEvent({ responseStatus: { response: 'declined' } })),
    ).toBe(false);
  });

  it('keeps an event marked free', () => {
    expect(isVisibleEvent(rawEvent({ showAs: 'free' }))).toBe(true);
  });

  it('keeps an event that has already finished', () => {
    expect(
      isVisibleEvent(
        rawEvent({
          start: { dateTime: '2026-09-12T02:00:00', timeZone: 'UTC' },
          end: { dateTime: '2026-09-12T03:00:00', timeZone: 'UTC' },
        }),
      ),
    ).toBe(true);
  });

  it('keeps an event whose responseStatus it does not recognise', () => {
    expect(
      isVisibleEvent(rawEvent({ responseStatus: { response: 'somethingNew' } })),
    ).toBe(true);
  });
});

describe('normaliseEvent', () => {
  it('maps every field the card renders', () => {
    const event = normaliseEvent(
      rawEvent({ isOnlineMeeting: true, type: 'occurrence', isOrganizer: true }),
    );
    expect(event).not.toBeNull();
    expect(event!.subject).toBe('Month-end close review');
    expect(event!.start.toISOString()).toBe('2026-09-12T09:00:00.000Z');
    expect(event!.end.toISOString()).toBe('2026-09-12T10:00:00.000Z');
    expect(event!.organiser).toBe('Rana Haddad');
    expect(event!.location).toBe('Meeting room 2');
    expect(event!.isOnlineMeeting).toBe(true);
    expect(event!.isRecurring).toBe(true);
    expect(event!.isOrganiser).toBe(true);
  });

  it('drops an event with no start rather than guessing one', () => {
    expect(normaliseEvent(rawEvent({ start: null }))).toBeNull();
    expect(normaliseEvents([rawEvent({ start: null }), rawEvent()])).toHaveLength(1);
  });
});

describe('eventOutlookUrl', () => {
  it('builds the current Outlook web link and encodes the id', () => {
    expect(eventOutlookUrl('AAMkAGI2=+/x', null)).toBe(
      'https://outlook.office365.com/calendar/item/AAMkAGI2%3D%2B%2Fx',
    );
  });

  it('falls back to webLink only when there is no id', () => {
    expect(eventOutlookUrl(null, 'https://old/link')).toBe('https://old/link');
    expect(eventOutlookUrl('id', 'https://old/link')).toContain('/calendar/item/id');
  });
});

describe('all-day events', () => {
  // Graph stores an all-day event at midnight with no meaningful zone, so the
  // value must not be converted, or it moves a day for any zone behind UTC.
  const onThe13th = normaliseEvent(
    rawEvent({
      isAllDay: true,
      subject: 'Company offsite',
      start: { dateTime: '2026-09-13T00:00:00.0000000', timeZone: 'UTC' },
      end: { dateTime: '2026-09-14T00:00:00.0000000', timeZone: 'UTC' },
    }),
  )!;

  it('lands on its stated day in a zone behind UTC', () => {
    // 13 Sep 00:00 UTC is 12 Sep 17:00 in Los Angeles. Converted, the event
    // would show as Today; floated, it correctly shows as Tomorrow.
    expect(bucketOf(onThe13th.start, NOW, LA, true)).toBe('tomorrow');
    expect(bucketOf(onThe13th.start, NOW, LA, false)).toBe('today');
  });

  it('lands on the same day in every zone', () => {
    for (const zone of [AMMAN, LA, 'UTC', 'Asia/Tokyo', 'Pacific/Kiritimati']) {
      expect(bucketOf(onThe13th.start, NOW, zone, true)).toBe('tomorrow');
    }
  });

  it('reaches the right column through groupItems, which sets the flag', () => {
    const grouped = groupItems([eventItem(onThe13th)], NOW, LA);
    expect(grouped.tomorrow).toHaveLength(1);
    expect(grouped.today).toHaveLength(0);
  });

  it('shows All day instead of a clock time', () => {
    expect(formatEventRange(onThe13th, LA)).toBe('All day');
  });
});

describe('formatEventRange', () => {
  it('renders a timed event in the viewer zone', () => {
    const event = normaliseEvent(rawEvent())!;
    expect(formatEventRange(event, AMMAN)).toBe('12:00 – 13:00');
    expect(formatEventRange(event, 'UTC')).toBe('09:00 – 10:00');
  });
});

describe('groupItems with a mixed day', () => {
  const task: HubTask = {
    id: 't1',
    title: 'D365 month-end checklist',
    listId: 'flagged',
    listName: 'Flagged email',
    isFlaggedEmail: true,
    // A task's due date sits at local midnight.
    due: new Date('2026-09-11T21:00:00Z'),
    status: 'notStarted',
    importance: 'normal',
    categories: [],
    bodyPreview: '',
    outlookUrl: null,
    sender: null,
    receivedDateTime: null,
    lastModified: null,
  };

  it('puts a task and a meeting on the same day in one column, task first', () => {
    const event = normaliseEvent(rawEvent())!;
    const grouped = groupItems([eventItem(event), taskItem(task)], NOW, AMMAN);
    expect(grouped.today).toHaveLength(2);
    expect(grouped.today[0].kind).toBe('task');
    expect(grouped.today[1].kind).toBe('event');
  });

  it('never puts an event in the no-due-date column', () => {
    const event = normaliseEvent(rawEvent())!;
    const grouped = groupItems([eventItem(event)], NOW, AMMAN);
    expect(grouped.noDueDate).toHaveLength(0);
  });
});
