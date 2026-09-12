import { describe, expect, it } from 'vitest';
import { applyTask, mergeWritten } from './mutate';
import type { HubTask } from './types';

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

describe('applyTask', () => {
  const original = task({
    id: 'a',
    title: 'Month-end close',
    due: new Date('2026-09-12T00:00:00Z'),
    sender: 'Rana Haddad',
    outlookUrl: 'https://outlook/a',
  });
  const others = [task({ id: 'b' }), task({ id: 'c' })];

  it('replaces only the matching task', () => {
    const next = applyTask([original, ...others], {
      ...original,
      status: 'completed',
    });
    expect(next[0].status).toBe('completed');
    expect(next[1]).toBe(others[0]);
    expect(next[2]).toBe(others[1]);
  });

  it('restores the previous task exactly when a write is rolled back', () => {
    const optimistic = { ...original, status: 'completed' };
    const afterWrite = applyTask([original, ...others], optimistic);
    const rolledBack = applyTask(afterWrite, original);
    expect(rolledBack[0]).toEqual(original);
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
