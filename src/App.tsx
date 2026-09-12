import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { IPublicClientApplication } from '@azure/msal-browser';
import { useAuth } from './auth/useAuth';
import { getTenantId, isConfigured } from './auth/msalConfig';
import { GraphClient } from './graph/client';
import { loadTasks } from './graph/todo';
import { completeTask, reopenTask, rescheduleTask } from './graph/write';
import { enrichTasks, fetchFlaggedMessages } from './graph/mail';
import { HORIZON_DAYS, fetchCalendarEvents } from './graph/calendar';
import {
  completePlannerTask,
  fetchPlannerTasks,
  reopenPlannerTask,
  reschedulePlannerTask,
} from './graph/planner';
import { SignInPanel } from './components/SignInPanel';
import { Timeline } from './components/Timeline';
import { viewerZone } from './model/dates';
import { listRefOf, normaliseTask } from './model/normalise';
import { applyItem, mergeWritten } from './model/mutate';
import { eventItem, itemKey, plannerItem, taskItem } from './model/types';
import type { HubItem, HubPlannerTask, HubTask } from './model/types';

interface Props {
  msal: IPublicClientApplication;
}

/** Which optional sources answered on the last read. */
interface Sources {
  senders: boolean;
  calendar: boolean;
  planner: boolean;
}

export function App({ msal }: Props) {
  const auth = useAuth(msal);
  const configured = isConfigured();
  const zone = viewerZone();

  const [items, setItems] = useState<HubItem[]>([]);
  const [loading, setLoading] = useState(false);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [sources, setSources] = useState<Sources>({
    senders: true,
    calendar: true,
    planner: true,
  });
  const [lastRefresh, setLastRefresh] = useState<Date | null>(null);
  const [now, setNow] = useState(() => new Date());

  const [pendingKeys, setPendingKeys] = useState<string[]>([]);
  const [undoItem, setUndoItem] = useState<HubItem | null>(null);
  // Guards against a second click landing on a write that is still in flight.
  const inFlight = useRef(new Set<string>());

  const refresh = useCallback(async () => {
    if (!auth.account) return;
    setLoading(true);
    setLoadError(null);
    try {
      const client = new GraphClient(auth.getToken);
      const todo = await loadTasks(client);

      // Senders, calendar and Planner are all optional. A failure in any of them
      // must not cost the timeline its flagged email.
      const readAt = new Date();
      const [messages, calendar, planner] = await Promise.all([
        fetchFlaggedMessages(client),
        fetchCalendarEvents(client, readAt, zone),
        fetchPlannerTasks(client, getTenantId()),
      ]);

      const tasks = messages ? enrichTasks(todo, messages) : todo;
      setSources({
        senders: messages !== null,
        calendar: calendar !== null,
        planner: planner !== null,
      });
      setItems([
        ...tasks.map(taskItem),
        ...(planner ?? []).map(plannerItem),
        ...(calendar ?? []).map(eventItem),
      ]);

      setUndoItem(null);
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
   * The optimistic item lands first, so the card moves column or leaves the
   * timeline on the click. A rejection restores the item exactly as it was: the
   * timeline must never show a change the service does not hold.
   *
   * `call` returns the item the service now holds, or null when the service
   * accepted the change but returned nothing worth rendering — a completed
   * Planner task, for instance, which the optimistic state already describes.
   */
  const mutate = useCallback(
    async (
      previous: HubItem,
      optimistic: HubItem,
      call: (client: GraphClient) => Promise<HubItem | null>,
    ): Promise<boolean> => {
      const key = itemKey(previous);
      if (inFlight.current.has(key)) return false;

      inFlight.current.add(key);
      setPendingKeys([...inFlight.current]);
      setItems((current) => applyItem(current, optimistic));
      setLoadError(null);

      try {
        const fresh = await call(new GraphClient(auth.getToken));
        if (fresh) setItems((current) => applyItem(current, fresh));
        return true;
      } catch (e) {
        setItems((current) => applyItem(current, previous));
        setLoadError(e instanceof Error ? e.message : String(e));
        return false;
      } finally {
        inFlight.current.delete(key);
        setPendingKeys([...inFlight.current]);
      }
    },
    [auth.getToken],
  );

  // Flagged email, through Microsoft To Do.

  const handleComplete = useCallback(
    async (task: HubTask) => {
      const previous = taskItem(task);
      const done = await mutate(
        previous,
        taskItem({ ...task, status: 'completed' }),
        async (client) =>
          taskItem(
            mergeWritten(
              task,
              normaliseTask(
                await completeTask(client, task.listId, task.id),
                listRefOf(task),
              ),
            ),
          ),
      );
      if (done) setUndoItem(previous);
    },
    [mutate],
  );

  const handleReschedule = useCallback(
    (task: HubTask, due: Date) => {
      void mutate(
        taskItem(task),
        taskItem({ ...task, due }),
        async (client) =>
          taskItem(
            mergeWritten(
              task,
              normaliseTask(
                await rescheduleTask(client, task.listId, task.id, due),
                listRefOf(task),
              ),
            ),
          ),
      );
    },
    [mutate],
  );

  // Planner.

  const handleCompletePlanner = useCallback(
    async (task: HubPlannerTask) => {
      const previous = plannerItem(task);
      const done = await mutate(
        previous,
        plannerItem({ ...task, percentComplete: 100 }),
        async (client) => {
          const fresh = await completePlannerTask(client, task, getTenantId());
          return fresh ? plannerItem(fresh) : null;
        },
      );
      if (done) setUndoItem(previous);
    },
    [mutate],
  );

  const handleReschedulePlanner = useCallback(
    (task: HubPlannerTask, due: Date) => {
      void mutate(
        plannerItem(task),
        plannerItem({ ...task, due }),
        async (client) => {
          const fresh = await reschedulePlannerTask(client, task, due, getTenantId());
          return fresh ? plannerItem(fresh) : null;
        },
      );
    },
    [mutate],
  );

  const handleUndo = useCallback(async () => {
    if (!undoItem) return;

    if (undoItem.kind === 'task') {
      const task = undoItem.task;
      const done = await mutate(
        taskItem({ ...task, status: 'completed' }),
        undoItem,
        async (client) =>
          taskItem(
            mergeWritten(
              task,
              normaliseTask(
                await reopenTask(client, task.listId, task.id),
                listRefOf(task),
              ),
            ),
          ),
      );
      if (done) setUndoItem(null);
      return;
    }

    if (undoItem.kind === 'planner') {
      const task = undoItem.task;
      const done = await mutate(
        plannerItem({ ...task, percentComplete: 100 }),
        undoItem,
        async (client) => {
          const fresh = await reopenPlannerTask(client, task, getTenantId());
          return fresh ? plannerItem(fresh) : undoItem;
        },
      );
      if (done) setUndoItem(null);
    }
  }, [mutate, undoItem]);

  const counts = useMemo(() => {
    let tasks = 0;
    let planner = 0;
    let events = 0;
    for (const item of items) {
      if (item.kind === 'task' && item.task.status !== 'completed') tasks += 1;
      if (item.kind === 'planner' && item.task.percentComplete < 100) planner += 1;
      if (item.kind === 'event') events += 1;
    }
    return { tasks, planner, events };
  }, [items]);

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

  const undoTitle =
    undoItem && undoItem.kind !== 'event' ? undoItem.task.title : null;
  const quiet = !loading && !loadError;

  return (
    <main className="shell">
      <header className="topbar">
        <div>
          <h1>AK MyToDo Hub</h1>
          <p className="lede">
            {counts.tasks} {counts.tasks === 1 ? 'task' : 'tasks'}
            {sources.planner && `, ${counts.planner} Planner`}
            {sources.calendar &&
              `, ${counts.events} ${counts.events === 1 ? 'event' : 'events'} to ${HORIZON_DAYS} days`}{' '}
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

      {!sources.senders && quiet && (
        <div className="notice">
          Sender names are unavailable for this mailbox. Flagged-email cards below
          show no sender.
        </div>
      )}

      {!sources.calendar && quiet && (
        <div className="notice">
          Calendar access is unavailable for this account. Add{' '}
          <code>Calendars.ReadBasic</code> to the app registration to see meetings.
        </div>
      )}

      {!sources.planner && quiet && (
        <div className="notice">
          Planner is unavailable for this account. The timeline shows Microsoft To
          Do and meetings only.
        </div>
      )}

      {undoTitle && (
        <div className="notice undo-bar">
          <span>
            Marked done: <strong>{undoTitle}</strong>
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
        pendingKeys={pendingKeys}
        onComplete={handleComplete}
        onReschedule={handleReschedule}
        onCompletePlanner={handleCompletePlanner}
        onReschedulePlanner={handleReschedulePlanner}
      />

      <footer className="foot">
        {lastRefresh
          ? `Last read ${lastRefresh.toLocaleTimeString('en-GB')}. Writes reach Microsoft To Do and Planner only.`
          : 'Writes reach Microsoft To Do and Planner only.'}
      </footer>
    </main>
  );
}
