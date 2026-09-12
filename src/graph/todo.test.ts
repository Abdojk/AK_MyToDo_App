import { describe, expect, it } from 'vitest';
import { loadTasks } from './todo';
import type { GraphClient } from './client';

/**
 * A stub standing in for GraphClient. loadTasks only reaches getAll and, when a
 * task arrives without linked resources, batchGet.
 */
function stubClient(pages: Record<string, unknown[]>): GraphClient {
  return {
    getAll: async (path: string) => {
      const key = Object.keys(pages).find((k) => path.startsWith(k));
      if (!key) throw new Error(`unexpected read: ${path}`);
      return pages[key];
    },
    batchGet: async () => new Map(),
  } as unknown as GraphClient;
}

describe('loadTasks', () => {
  const client = stubClient({
    '/me/todo/lists?': [
      { id: 'L1', displayName: 'Flagged email', wellknownListName: 'flaggedEmails' },
      { id: 'L2', displayName: 'Tasks', wellknownListName: 'defaultList' },
      { id: 'L3', displayName: 'Engicon rollout', wellknownListName: 'none' },
    ],
    '/me/todo/lists/L1/tasks': [
      { id: 't1', title: 'RE: Month-end close', linkedResources: [] },
    ],
    '/me/todo/lists/L2/tasks': [{ id: 't2', title: 'Book the flights', linkedResources: [] }],
    '/me/todo/lists/L3/tasks': [{ id: 't3', title: 'Draft the SOW', linkedResources: [] }],
  });

  it('reads every list, not only the flagged one', async () => {
    const tasks = await loadTasks(client);
    expect(tasks.map((t) => t.title)).toEqual([
      'RE: Month-end close',
      'Book the flights',
      'Draft the SOW',
    ]);
  });

  it('gives each task the list it actually lives in, which its writes need', async () => {
    const tasks = await loadTasks(client);
    expect(tasks.map((t) => t.listId)).toEqual(['L1', 'L2', 'L3']);
    expect(tasks.map((t) => t.listName)).toEqual([
      'Flagged email',
      'Tasks',
      'Engicon rollout',
    ]);
  });

  it('marks only the built-in Flagged email list', async () => {
    const tasks = await loadTasks(client);
    expect(tasks.map((t) => t.isFlaggedEmail)).toEqual([true, false, false]);
  });

  it('does not fail a mailbox with no flagged list', async () => {
    const noFlagged = stubClient({
      '/me/todo/lists?': [{ id: 'L2', displayName: 'Tasks', wellknownListName: 'defaultList' }],
      '/me/todo/lists/L2/tasks': [{ id: 't2', title: 'Book the flights', linkedResources: [] }],
    });
    await expect(loadTasks(noFlagged)).resolves.toHaveLength(1);
  });

  it('names an unnamed list rather than rendering a blank', async () => {
    const unnamed = stubClient({
      '/me/todo/lists?': [{ id: 'L9', displayName: '  ' }],
      '/me/todo/lists/L9/tasks': [{ id: 't9', title: 'Something', linkedResources: [] }],
    });
    const [task] = await loadTasks(unnamed);
    expect(task.listName).toBe('Tasks');
  });
});
