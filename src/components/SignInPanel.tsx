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
        Every flagged Outlook email, laid out by due date.
      </p>

      {!configured && (
        <div className="notice notice-error">
          <strong>Configuration missing.</strong> Copy <code>.env.example</code> to{' '}
          <code>.env</code> and fill in <code>VITE_CLIENT_ID</code> and{' '}
          <code>VITE_TENANT_ID</code> from the Entra ID app registration. The
          README covers the registration steps.
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
        The Hub requests <code>User.Read</code>, <code>Tasks.ReadWrite</code>, and{' '}
        <code>Mail.ReadBasic</code>. It writes only to your Microsoft To Do tasks,
        to mark one done or move its due date. Mail access stays read-only.
      </p>
    </div>
  );
}
