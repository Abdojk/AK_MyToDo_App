import { useCallback, useEffect, useState } from 'react';
import type { AccountInfo, IPublicClientApplication } from '@azure/msal-browser';
import { InteractionRequiredAuthError } from '@azure/msal-browser';
import { graphScopes, loginRequest } from './msalConfig';

export interface AuthState {
  account: AccountInfo | null;
  signIn: () => Promise<void>;
  signOut: () => Promise<void>;
  getToken: () => Promise<string>;
  busy: boolean;
  error: string | null;
}

/**
 * Wraps MSAL in the three calls the Hub needs: sign in, sign out, and fetch a
 * Graph access token. Token renewal stays silent unless Entra ID asks for
 * interaction, in which case a popup handles it.
 */
export function useAuth(msal: IPublicClientApplication): AuthState {
  const [account, setAccount] = useState<AccountInfo | null>(
    () => msal.getActiveAccount() ?? msal.getAllAccounts()[0] ?? null,
  );
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    if (account && !msal.getActiveAccount()) msal.setActiveAccount(account);
  }, [account, msal]);

  const signIn = useCallback(async () => {
    setBusy(true);
    setError(null);
    try {
      const result = await msal.loginPopup(loginRequest);
      msal.setActiveAccount(result.account);
      setAccount(result.account);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }, [msal]);

  const signOut = useCallback(async () => {
    setBusy(true);
    try {
      await msal.logoutPopup({ account: account ?? undefined });
      setAccount(null);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  }, [msal, account]);

  const getToken = useCallback(async (): Promise<string> => {
    const target = msal.getActiveAccount() ?? account;
    if (!target) throw new Error('Sign in before calling Microsoft Graph.');
    try {
      const result = await msal.acquireTokenSilent({
        scopes: graphScopes,
        account: target,
      });
      return result.accessToken;
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        const result = await msal.acquireTokenPopup({ scopes: graphScopes });
        return result.accessToken;
      }
      throw e;
    }
  }, [msal, account]);

  return { account, signIn, signOut, getToken, busy, error };
}
