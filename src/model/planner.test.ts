import { describe, expect, it } from 'vitest';
import {
  isPlannerComplete,
  isPressing,
  normalisePlannerTask,
  plannerPriority,
  plannerTaskUrl,
  toPlannerDueDate,
} from './planner';
import { groupItems } from './buckets';
import { eventItem, plannerItem, taskItem } from './types';
import type { HubTask, RawPlannerTask } from './types';
import { normaliseEvent } from './events';

const TENANT = 'contoso.onmicrosoft.com';
const AMMAN = 'Asia/Amman';
const NOW = new Date('2026-09-12T07:00:00Z');
const NO_PLANS = new Map<string, string>();

const raw = (over: Partial<RawPlannerTask> = {}): RawPlannerTask => ({
  id: '01gzSlKkIUSUl6DF_EilrmQAKDhh',
  title: 'Migrate vendor master data',
  planId: 'xqQg5FS2LkCp935s-FIFm2QAFkHM',
  dueDateTime: '2026-09-16T00:00:00Z',
  percentComplete: 50,
  priority: 3,
  checklistItemCount: 5,
  activeChecklistItemCount: 3,
  '@odata.etag': 'W/"JzEtVGFzayc="',
  ...over,
});

describe('plannerPriority', () => {
  // Planner reads 0-1 as urgent, 2-4 important, 5-7 medium, 8-10 low.
  it('maps every band at its boundaries', () => {
    expect([0, 1].map(plannerPriority)).toEqual(['urgent', 'urgent']);
    expect([2, 4].map(plannerPriority)).toEqual(['important', 'important']);
    expect([5, 7].map(plannerPriority)).toEqual(['medium', 'medium']);
    expect([8, 10].map(plannerPriority)).toEqual(['low', 'low']);
  });

  it('treats an absent priority as medium', () => {
    expect(plannerPriority(null)).toBe('medium');
    expect(plannerPriority(undefined)).toBe('medium');
  });

  it('marks only urgent and important as pressing', () => {
    expect((['urgent', 'important'] as const).map(isPressing)).toEqual([true, true]);
    expect((['medium', 'low'] as const).map(isPressing)).toEqual([false, false]);
  });
});

describe('normalisePlannerTask', () => {
  it('maps every field the card renders', () => {
    const task = normalisePlannerTask(
      raw(),
      new Map([['xqQg5FS2LkCp935s-FIFm2QAFkHM', 'Engicon rollout']]),
      TENANT,
    );
    expect(task).not.toBeNull();
    expect(task!.title).toBe('Migrate vendor master data');
    expect(task!.due.toISOString()).toBe('2026-09-16T00:00:00.000Z');
    expect(task!.priority).toBe('important');
    expect(task!.planTitle).toBe('Engicon rollout');
    expect(task!.checklistDone).toBe(2);
    expect(task!.checklistTotal).toBe(5);
    expect(task!.etag).toBe('W/"JzEtVGFzayc="');
  });

  it('drops a task with no due date', () => {
    expect(normalisePlannerTask(raw({ dueDateTime: null }), NO_PLANS, TENANT)).toBeNull();
  });

  it('drops a finished task', () => {
    expect(isPlannerComplete(raw({ percentComplete: 100 }))).toBe(true);
    expect(
      normalisePlannerTask(raw({ percentComplete: 100 }), NO_PLANS, TENANT),
    ).toBeNull();
  });

  it('survives a task with no plan title, checklist, or etag', () => {
    const task = normalisePlannerTask(
      raw({ checklistItemCount: null, activeChecklistItemCount: null, '@odata.etag': null }),
      NO_PLANS,
      TENANT,
    );
    expect(task!.planTitle).toBeNull();
    expect(task!.checklistTotal).toBe(0);
    expect(task!.etag).toBeNull();
  });
});

describe('plannerTaskUrl', () => {
  it('encodes the tenant and the task id', () => {
    expect(plannerTaskUrl(TENANT, '01gz_Sl-Kk')).toBe(
      'https://tasks.office.com/contoso.onmicrosoft.com/Home/Task/01gz_Sl-Kk',
    );
  });

  it('returns nothing when the tenant is unset, rather than a broken link', () => {
    expect(plannerTaskUrl(undefined, 'x')).toBe('');
  });
});

describe('toPlannerDueDate', () => {
  // Planner takes a bare DateTimeOffset, not the dateTimeTimeZone pair todoTask
  // uses, so the write format differs from src/model/reschedule.ts.
  it('emits a UTC ISO instant', () => {
    expect(toPlannerDueDate(new Date('2026-09-19T21:00:00Z'))).toBe(
      '2026-09-19T21:00:00.000Z',
    );
  });
});

describe('groupItems with all three sources', () => {
  const todo: HubTask = {
    id: 't1',
    title: 'D365 month-end checklist',
    due: new Date('2026-09-11T21:00:00Z'), // local midnight on the 12th in Amman
    status: 'notStarted',
    importance: 'normal',
    categories: [],
    bodyPreview: '',
    outlookUrl: null,
    sender: null,
    receivedDateTime: null,
    lastModified: null,
  };

  const meeting = normaliseEvent({
    id: 'e1',
    subject: 'Month-end close review',
    start: { dateTime: '2026-09-12T09:00:00', timeZone: 'UTC' },
    end: { dateTime: '2026-09-12T10:00:00', timeZone: 'UTC' },
    isAllDay: false,
    isCancelled: false,
    type: 'singleInstance',
  })!;

  const planner = normalisePlannerTask(
    raw({ dueDateTime: '2026-09-12T05:00:00Z' }),
    NO_PLANS,
    TENANT,
  )!;

  it('puts all three on the same day in one column, tasks before the meeting', () => {
    const grouped = groupItems(
      [eventItem(meeting), plannerItem(planner), taskItem(todo)],
      NOW,
      AMMAN,
    );
    expect(grouped.today.map((i) => i.kind)).toEqual(['task', 'planner', 'event']);
  });

  it('drops a Planner task once it reaches 100 per cent', () => {
    const grouped = groupItems(
      [plannerItem({ ...planner, percentComplete: 100 })],
      NOW,
      AMMAN,
    );
    expect(grouped.today).toHaveLength(0);
  });
});
