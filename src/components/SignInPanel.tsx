interface Props {
  onSignIn: () => void;
  busy: boolean;
  error: string | null;
  configured: boolean;
}

export function SignInPanel({ onSignIn, busy, error, configured }: Props) {
  return (
    <div className="panel">
      <h1>AK MyToDo Hub</h1>
      <p className="lede">
        Your Microsoft To Do tasks, Planner tasks and meetings, laid out by date.
      </p>

      {!configured && (
        <div className="notice notice-error">
          <strong>Configuration missing.</strong> Running it yourself: use{' '}
          <code>start.cmd</code>, which asks for the two IDs. Hosted: put them in{' '}
          <code>config.json</code> beside this page. Both come from the Entra ID
          app registration, which the README covers.
        </div>
      )}

      <button
        type="button"
        className="primary"
        onClick={onSignIn}
        disabled={busy || !configured}
      >
        {busy ? 'Signing in…' : 'Sign in with Microsoft'}
      </button>

      {error && <div className="notice notice-error">{error}</div>}

      <p className="fine">
        The Hub requests <code>User.Read</code>, <code>Tasks.ReadWrite</code>,{' '}
        <code>Mail.ReadBasic</code> and <code>Calendars.ReadBasic</code>. It writes
        only to your Microsoft To Do and Planner tasks, to mark one done or move
        its due date. Mail and calendar access stay read-only.
      </p>
    </div>
  );
}
