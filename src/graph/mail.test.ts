import { describe, expect, it } from 'vitest';
import { enrichTasks } from './mail';
import type { HubTask } from '../model/types';

const task = (id: string, title: string): HubTask => ({
  id,
  title,
  due: null,
  status: 'notStarted',
  importance: 'normal',
  categories: [],
  bodyPreview: '',
  outlookUrl: null,
  sender: null,
  receivedDateTime: null,
  lastModified: null,
});

describe('enrichTasks', () => {
  it('attaches the sender when the subject matches once', () => {
    const [result] = enrichTasks(
      [task('1', 'RE: Month-end close')],
      [
        {
          id: 'm1',
          subject: 'Month-end close',
          receivedDateTime: '2026-09-10T08:00:00Z',
          webLink: 'https://outlook.office.com/mail/m1',
          from: { emailAddress: { name: 'Rana Haddad', address: 'rana@example.com' } },
        },
      ],
    );
    expect(result.sender).toBe('Rana Haddad');
    expect(result.outlookUrl).toBe('https://outlook.office.com/mail/m1');
  });

  it('leaves the sender null when the subject is ambiguous', () => {
    const [result] = enrichTasks(
      [task('1', 'Status update')],
      [
        { id: 'a', subject: 'Status update', from: { emailAddress: { name: 'A' } } },
        { id: 'b', subject: 'RE: Status update', from: { emailAddress: { name: 'B' } } },
      ],
    );
    expect(result.sender).toBeNull();
  });

  it('keeps a linked-resource URL in preference to the message link', () => {
    const withUrl = { ...task('1', 'Invoice'), outlookUrl: 'https://linked/1' };
    const [result] = enrichTasks(
      [withUrl],
      [{ id: 'm', subject: 'Invoice', webLink: 'https://message/1' }],
    );
    expect(result.outlookUrl).toBe('https://linked/1');
  });
});
