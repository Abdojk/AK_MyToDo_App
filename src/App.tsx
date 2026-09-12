import { useCallback, useEffect, useRef, useState } from 'react';
import type { IPublicClientApplication } from '@azure/msal-browser';
import { useAuth } from './auth/useAuth';
import { isConfigured } from './auth/msalConfig';
import { GraphClient } from './graph/client';
import { loadFlaggedTasks } from './graph/todo';
import { completeTask, reopenTask, rescheduleTask } from './graph/write';
import { enrichTasks, fetchFlaggedMessages } from './graph/mail';
import { HORIZON_DAYS, fetchCalendarEvents } from './graph/calendar';
import { SignInPanel } from './components/SignInPanel';
import { Timeline } from './components/Timeline';
import { viewerZone } from './model/dates';
import { normaliseTask } from './model/normalise';
import { applyTask, mergeWritten } from './model/mutate';
import { eventItem, taskItem } from './model/types';
import type { HubEvent, HubItem, HubTask, RawTodoTask } from './model/types';

interface Props {
  msal: IPublicClientApplication;
}

export function App({ msal }: Props) {
  const auth = useAuth(msal);
  const configured = isConfigured();
  const zone = viewerZone();

  const [listId, setListId] = useState<string | null>(null);
  const [tasks, setTasks] = useState<HubTask[]>([]);
  const [loading, setLoading] = useState(false);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [events, setEvents] = useState<HubEvent[]>([]);
  const [enriched, setEnriched] = useState(false);
  const [calendarOn, setCalendarOn] = useState(true);
  const [lastRefresh, setLastRefresh] = useState<Date | null>(null);
  const [now, setNow] = useState(() => new Date());

  const [pendingIds, setPendingIds] = useState<string[]>([]);
  const [undoTask, setUndoTask] = useState<HubTask | null>(null);
  // Guards against a second click landing on a write that is still in flight.
  const inFlight = useRef(new Set<string>());

  const refresh = useCallback(async () => {
    if (!auth.account) return;
    setLoading(true);
    setLoadError(null);
    try {
      const client = new GraphClient(auth.getToken);
      const { listId: flaggedListId, tasks: flagged } = await loadFlaggedTasks(client);
      setListId(flaggedListId);

      // Sender names and the calendar are both optional. A failure in either
      // must not cost the timeline its tasks.
      const readAt = new Date();
      const [messages, calendar] = await Promise.all([
        fetchFlaggedMessages(client),
        fetchCalendarEvents(client, readAt, zone),
      ]);
      setEnriched(messages !== null);
      setTasks(messages ? enrichTasks(flagged, messages) : flagged);
      setCalendarOn(calendar !== null);
      setEvents(calendar ?? []);

      setUndoTask(null);
      setNow(new Date());
      setLastRefresh(new Date());
    } catch (e) {
      setLoadError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }, [auth.account, auth.getToken, zone]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  // Roll the buckets over at midnight without a manual reload.
  useEffect(() => {
    const timer = window.setInterval(() => setNow(new Date()), 60_000);
    return () => window.clearInterval(timer);
  }, []);

  /**
   * Apply a write.
   *
   * The optimistic task lands first, so the card moves column or leaves the
   * timeline on the click. A rejection restores the task exactly as it was: the
   * timeline must never show a change the mailbox does not hold.
   */
  const mutate = useCallback(
    async (
      previous: HubTask,
      optimistic: HubTask,
      call: (client: GraphClient, listId: string) => Promise<RawTodoTask>,
    ): Promise<boolean> => {
      if (!listId || inFlight.current.has(previous.id)) return false;

      inFlight.current.add(previous.id);
      setPendingIds([...inFlight.current]);
      setTasks((current) => applyTask(current, optimistic));
      setLoadError(null);

      try {
        const raw = await call(new GraphClient(auth.getToken), listId);
        setTasks((current) =>
          applyTask(current, mergeWritten(previous, normaliseTask(raw, []))),
        );
        return true;
      } catch (e) {
        setTasks((current) => applyTask(current, previous));
        setLoadError(e instanceof Error ? e.message : String(e));
        return false;
      } finally {
        inFlight.current.delete(previous.id);
        setPendingIds([...inFlight.current]);
      }
    },
    [listId, auth.getToken],
  );

  const handleComplete = useCallback(
    async (task: HubTask) => {
      const done = await mutate(
        task,
        { ...task, status: 'completed' },
        (client, id) => completeTask(client, id, task.id),
      );
      if (done) setUndoTask(task);
    },
    [mutate],
  );

  const handleUndo = useCallback(async () => {
    if (!undoTask) return;
    const done = await mutate(
      { ...undoTask, status: 'completed' },
      undoTask,
      (client, id) => reopenTask(client, id, undoTask.id),
    );
    if (done) setUndoTask(null);
  }, [mutate, undoTask]);

  const handleReschedule = useCallback(
    (task: HubTask, due: Date) =>
      mutate(task, { ...task, due }, (client, id) =>
        rescheduleTask(client, id, task.id, due),
      ).then(() => undefined),
    [mutate],
  );

  if (!auth.account) {
    return (
      <main className="shell shell-centred">
        <SignInPanel
          onSignIn={() => void auth.signIn()}
          busy={auth.busy}
          error={auth.error}
          configured={configured}
        />
      </main>
    );
  }

  const openCount = tasks.filter((t) => t.status !== 'completed').length;
  const items: HubItem[] = [...tasks.map(taskItem), ...events.map(eventItem)];

  return (
    <main className="shell">
      <header className="topbar">
        <div>
          <h1>AK MyToDo Hub</h1>
          <p className="lede">
            {openCount} flagged {openCount === 1 ? 'email' : 'emails'}
            {calendarOn &&
              `, ${events.length} ${events.length === 1 ? 'event' : 'events'} to ${HORIZON_DAYS} days`}{' '}
            · {auth.account.username} · {zone}
          </p>
        </div>
        <div className="topbar-actions">
          <button type="button" onClick={() => void refresh()} disabled={loading}>
            {loading ? 'Refreshing…' : 'Refresh'}
          </button>
          <button type="button" className="ghost" onClick={() => void auth.signOut()}>
            Sign out
          </button>
        </div>
      </header>

      {loadError && <div className="notice notice-error">{loadError}</div>}

      {!enriched && !loading && !loadError && (
        <div className="notice">
          Sender names are unavailable for this mailbox. The timeline below reads
          from Microsoft To Do alone.
        </div>
      )}

      {!calendarOn && !loading && !loadError && (
        <div className="notice">
          Calendar access is unavailable for this account. The timeline below
          shows tasks only. Add <code>Calendars.ReadBasic</code> to the app
          registration to see meetings.
        </div>
      )}

      {undoTask && (
        <div className="notice undo-bar">
          <span>
            Marked done: <strong>{undoTask.title}</strong>
          </span>
          <button type="button" onClick={() => void handleUndo()}>
            Undo
          </button>
        </div>
      )}

      <Timeline
        items={items}
        now={now}
        zone={zone}
        pendingIds={pendingIds}
        onComplete={handleComplete}
        onReschedule={handleReschedule}
      />

      <footer className="foot">
        {lastRefresh
          ? `Last read ${lastRefresh.toLocaleTimeString('en-GB')}. Writes reach Microsoft To Do only.`
          : 'Writes reach Microsoft To Do only.'}
      </footer>
    </main>
  );
}
