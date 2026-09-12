import { describe, expect, it } from 'vitest';
import { bucketOf, groupItems } from './buckets';
import { resolveGraphDateTime } from './dates';
import { normaliseTask } from './normalise';
import { taskItem } from './types';
import type { HubTask, RawTodoTask } from './types';

const ZONE = 'Asia/Amman';
// 12 Sep 2026, 10:00 in Asia/Amman (UTC+3 in September).
const NOW = new Date('2026-09-12T07:00:00Z');

function dueAt(local: string): Date {
  return resolveGraphDateTime({ dateTime: local, timeZone: ZONE })!;
}

describe('bucketOf', () => {
  it('puts a missing due date in noDueDate', () => {
    expect(bucketOf(null, NOW, ZONE)).toBe('noDueDate');
  });

  it('puts yesterday in overdue', () => {
    expect(bucketOf(dueAt('2026-09-11T09:00:00'), NOW, ZONE)).toBe('overdue');
  });

  it('puts one minute past midnight today in today', () => {
    expect(bucketOf(dueAt('2026-09-12T00:01:00'), NOW, ZONE)).toBe('today');
  });

  it('puts one minute before midnight today in today, not tomorrow', () => {
    expect(bucketOf(dueAt('2026-09-12T23:59:00'), NOW, ZONE)).toBe('today');
  });

  it('puts tomorrow in tomorrow', () => {
    expect(bucketOf(dueAt('2026-09-13T08:00:00'), NOW, ZONE)).toBe('tomorrow');
  });

  it('puts day two through day seven in next7', () => {
    expect(bucketOf(dueAt('2026-09-14T08:00:00'), NOW, ZONE)).toBe('next7');
    expect(bucketOf(dueAt('2026-09-19T08:00:00'), NOW, ZONE)).toBe('next7');
  });

  it('puts day eight in later', () => {
    expect(bucketOf(dueAt('2026-09-20T08:00:00'), NOW, ZONE)).toBe('later');
  });

  it('reads a due date stated in a different zone as the same instant', () => {
    // 13 Sep 00:30 in Tokyo is still 12 Sep 18:30 in Amman, so this is today.
    const tokyo = resolveGraphDateTime({
      dateTime: '2026-09-13T00:30:00',
      timeZone: 'Asia/Tokyo',
    })!;
    expect(bucketOf(tokyo, NOW, ZONE)).toBe('today');
  });

  it('falls back to UTC for a zone name Intl does not accept', () => {
    const windowsZone = resolveGraphDateTime({
      dateTime: '2026-09-12T22:00:00',
      timeZone: 'Pacific Standard Time',
    })!;
    expect(windowsZone.toISOString()).toBe('2026-09-12T22:00:00.000Z');
  });
});

describe('groupItems', () => {
  const base: HubTask = {
    id: 'x',
    title: 'Task',
    listId: 'flagged',
    listName: 'Flagged email',
    isFlaggedEmail: true,
    due: null,
    status: 'notStarted',
    importance: 'normal',
    categories: [],
    bodyPreview: '',
    outlookUrl: null,
    sender: null,
    receivedDateTime: null,
    lastModified: null,
  };

  it('drops completed tasks', () => {
    const grouped = groupItems(
      [taskItem({ ...base, id: 'a', status: 'completed', due: dueAt('2026-09-12T09:00:00') })],
      NOW,
      ZONE,
    );
    expect(grouped.today).toHaveLength(0);
  });

  it('sorts a column by due date, then importance', () => {
    const grouped = groupItems(
      [
        taskItem({ ...base, id: 'late', due: dueAt('2026-09-12T17:00:00') }),
        taskItem({ ...base, id: 'early', due: dueAt('2026-09-12T08:00:00') }),
        taskItem({
          ...base,
          id: 'early-high',
          due: dueAt('2026-09-12T08:00:00'),
          importance: 'high',
        }),
      ],
      NOW,
      ZONE,
    );
    expect(grouped.today.map((i) => (i.kind === 'task' ? i.task.id : ''))).toEqual([
      'early-high',
      'early',
      'late',
    ]);
  });
});

const LIST = { id: 'flagged', displayName: 'Flagged email', isFlaggedEmail: true };

describe('normaliseTask', () => {
  const raw: RawTodoTask = {
    id: 'AAMkADIyAAAhrbPWAAA=',
    title: 'RE: D365 F&O month-end close',
    body: { content: '<p>Needs the <b>GL</b> reconciliation</p>', contentType: 'html' },
    status: 'notStarted',
    importance: 'high',
    categories: ['Finance'],
    dueDateTime: { dateTime: '2026-09-15T00:00:00.0000000', timeZone: 'UTC' },
    lastModifiedDateTime: '2026-09-10T11:22:33Z',
    linkedResources: [
      {
        id: 'lr1',
        applicationName: 'Microsoft Outlook',
        displayName: 'RE: D365 F&O month-end close',
        externalId: 'abc',
        webUrl: 'https://outlook.office.com/mail/deeplink/abc',
      },
    ],
  };

  it('maps every field the timeline renders', () => {
    const task = normaliseTask(raw, LIST);
    expect(task.id).toBe(raw.id);
    expect(task.title).toBe('RE: D365 F&O month-end close');
    expect(task.due?.toISOString()).toBe('2026-09-15T00:00:00.000Z');
    expect(task.importance).toBe('high');
    expect(task.categories).toEqual(['Finance']);
    expect(task.bodyPreview).toBe('Needs the GL reconciliation');
    expect(task.outlookUrl).toBe('https://outlook.office.com/mail/deeplink/abc');
    expect(task.lastModified?.toISOString()).toBe('2026-09-10T11:22:33.000Z');
  });

  it('carries the list it was read from onto the task', () => {
    const custom = { id: 'AAMk2=', displayName: 'Engicon', isFlaggedEmail: false };
    const task = normaliseTask(raw, custom);
    expect(task.listId).toBe('AAMk2=');
    expect(task.listName).toBe('Engicon');
    expect(task.isFlaggedEmail).toBe(false);
  });

  it('survives a task with no due date, body, or linked resource', () => {
    const task = normaliseTask({ id: 'bare' }, LIST);
    expect(task.due).toBeNull();
    expect(task.title).toBe('(no subject)');
    expect(task.outlookUrl).toBeNull();
    expect(task.bodyPreview).toBe('');
  });
});
