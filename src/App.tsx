import { useCallback, useEffect, useState } from 'react';
import type { IPublicClientApplication } from '@azure/msal-browser';
import { useAuth } from './auth/useAuth';
import { isConfigured } from './auth/msalConfig';
import { GraphClient } from './graph/client';
import { loadFlaggedTasks } from './graph/todo';
import { enrichTasks, fetchFlaggedMessages } from './graph/mail';
import { SignInPanel } from './components/SignInPanel';
import { Timeline } from './components/Timeline';
import { viewerZone } from './model/dates';
import type { HubTask } from './model/types';

interface Props {
  msal: IPublicClientApplication;
}

export function App({ msal }: Props) {
  const auth = useAuth(msal);
  const configured = isConfigured();
  const zone = viewerZone();

  const [tasks, setTasks] = useState<HubTask[]>([]);
  const [loading, setLoading] = useState(false);
  const [loadError, setLoadError] = useState<string | null>(null);
  const [enriched, setEnriched] = useState(false);
  const [lastRefresh, setLastRefresh] = useState<Date | null>(null);
  const [now, setNow] = useState(() => new Date());

  const refresh = useCallback(async () => {
    if (!auth.account) return;
    setLoading(true);
    setLoadError(null);
    try {
      const client = new GraphClient(auth.getToken);
      const flagged = await loadFlaggedTasks(client);

      // Sender names are a bonus. A failure here must not cost the timeline.
      const messages = await fetchFlaggedMessages(client);
      setEnriched(messages !== null);
      setTasks(messages ? enrichTasks(flagged, messages) : flagged);

      setNow(new Date());
      setLastRefresh(new Date());
    } catch (e) {
      setLoadError(e instanceof Error ? e.message : String(e));
    } finally {
      setLoading(false);
    }
  }, [auth.account, auth.getToken]);

  useEffect(() => {
    void refresh();
  }, [refresh]);

  // Roll the buckets over at midnight without a manual reload.
  useEffect(() => {
    const timer = window.setInterval(() => setNow(new Date()), 60_000);
    return () => window.clearInterval(timer);
  }, []);

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

  return (
    <main className="shell">
      <header className="topbar">
        <div>
          <h1>AK MyToDo Hub</h1>
          <p className="lede">
            {openCount} flagged {openCount === 1 ? 'email' : 'emails'} ·{' '}
            {auth.account.username} · {zone}
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

      <Timeline tasks={tasks} now={now} zone={zone} />

      <footer className="foot">
        {lastRefresh
          ? `Last read ${lastRefresh.toLocaleTimeString('en-GB')}. Read-only: the Hub never writes to the mailbox.`
          : 'Read-only: the Hub never writes to the mailbox.'}
      </footer>
    </main>
  );
}
