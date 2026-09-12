import { describe, expect, it } from 'vitest';
import { applyItem, mergeWritten } from './mutate';
import { plannerItem, taskItem } from './types';
import type { HubPlannerTask, HubTask } from './types';

const task = (over: Partial<HubTask> & { id: string }): HubTask => ({
  title: 'Task',
  due: null,
  status: 'notStarted',
  importance: 'normal',
  categories: [],
  bodyPreview: '',
  outlookUrl: null,
  sender: null,
  receivedDateTime: null,
  lastModified: null,
  ...over,
});

const plannerTask = (over: Partial<HubPlannerTask> & { id: string }): HubPlannerTask => ({
  title: 'Planner task',
  due: new Date('2026-09-16T00:00:00Z'),
  percentComplete: 0,
  priority: 'medium',
  planId: 'plan1',
  planTitle: 'Engicon rollout',
  checklistDone: 0,
  checklistTotal: 0,
  etag: 'W/"abc"',
  url: 'https://tasks.office.com/t/Home/Task/p1',
  ...over,
});

describe('applyItem', () => {
  const original = task({
    id: 'a',
    title: 'Month-end close',
    due: new Date('2026-09-12T00:00:00Z'),
    sender: 'Rana Haddad',
    outlookUrl: 'https://outlook/a',
  });
  const others = [taskItem(task({ id: 'b' })), plannerItem(plannerTask({ id: 'a' }))];

  it('replaces only the matching item', () => {
    const next = applyItem([taskItem(original), ...others], {
      kind: 'task',
      task: { ...original, status: 'completed' },
    });
    expect(next[0].kind === 'task' && next[0].task.status).toBe('completed');
    expect(next[1]).toBe(others[0]);
    expect(next[2]).toBe(others[1]);
  });

  it('keeps sources apart when a To Do task and a Planner task share an id', () => {
    // Both carry id 'a'. Writing the Planner one must not touch the To Do one.
    const next = applyItem(
      [taskItem(original), ...others],
      plannerItem(plannerTask({ id: 'a', percentComplete: 100 })),
    );
    expect(next[0].kind === 'task' && next[0].task.status).toBe('notStarted');
    expect(next[2].kind === 'planner' && next[2].task.percentComplete).toBe(100);
  });

  it('restores the previous item exactly when a write is rolled back', () => {
    const optimistic = { ...original, status: 'completed' };
    const afterWrite = applyItem([taskItem(original), ...others], taskItem(optimistic));
    const rolledBack = applyItem(afterWrite, taskItem(original));
    expect(rolledBack[0]).toEqual(taskItem(original));
    expect(rolledBack).toHaveLength(3);
  });
});

describe('mergeWritten', () => {
  it('keeps the fields a PATCH response cannot carry', () => {
    const previous = task({
      id: 'a',
      sender: 'Omar Nasser',
      outlookUrl: 'https://outlook/a',
      receivedDateTime: new Date('2026-09-10T08:00:00Z'),
    });
    const fresh = task({
      id: 'a',
      status: 'completed',
      due: new Date('2026-09-19T00:00:00Z'),
    });

    const merged = mergeWritten(previous, fresh);
    expect(merged.status).toBe('completed');
    expect(merged.due).toEqual(fresh.due);
    expect(merged.sender).toBe('Omar Nasser');
    expect(merged.outlookUrl).toBe('https://outlook/a');
    expect(merged.receivedDateTime).toEqual(previous.receivedDateTime);
  });
});
