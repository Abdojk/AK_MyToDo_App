import type { Configuration, PopupRequest } from '@azure/msal-browser';
import { LogLevel } from '@azure/msal-browser';

function required(name: string, value: string | undefined): string {
  if (!value || value.startsWith('00000000-')) {
    throw new Error(
      `${name} is not set. Copy .env.example to .env and fill in the values from the Entra ID app registration (see README).`,
    );
  }
  return value;
}

export const clientId = import.meta.env.VITE_CLIENT_ID as string | undefined;
export const tenantId = import.meta.env.VITE_TENANT_ID as string | undefined;

/** Resolved lazily: `window` does not exist when the module loads under a test runner. */
export function resolveRedirectUri(): string {
  const configured = import.meta.env.VITE_REDIRECT_URI as string | undefined;
  if (configured) return configured;
  return typeof window === 'undefined' ? 'http://localhost:5173' : window.location.origin;
}

/** True once .env carries a real client and tenant ID. */
export function isConfigured(): boolean {
  try {
    required('VITE_CLIENT_ID', clientId);
    required('VITE_TENANT_ID', tenantId);
    return true;
  } catch {
    return false;
  }
}

export function buildMsalConfig(): Configuration {
  return {
    auth: {
      clientId: required('VITE_CLIENT_ID', clientId),
      authority: `https://login.microsoftonline.com/${required('VITE_TENANT_ID', tenantId)}`,
      redirectUri: resolveRedirectUri(),
      postLogoutRedirectUri: resolveRedirectUri(),
      navigateToLoginRequestUrl: false,
    },
    cache: {
      // Session storage clears when the tab closes, which suits a single-user
      // read-only Hub better than persisting tokens across browser restarts.
      cacheLocation: 'sessionStorage',
      storeAuthStateInCookie: false,
    },
    system: {
      loggerOptions: {
        logLevel: LogLevel.Warning,
        loggerCallback: (_level, message, containsPii) => {
          if (!containsPii) console.warn('[msal]', message);
        },
      },
    },
  };
}

/**
 * Read-only delegated scopes.
 *
 * Mail.ReadBasic backs the optional sender enrichment. Remove it here and from
 * the app registration to run the Hub on Microsoft To Do data alone.
 */
export const graphScopes = ['User.Read', 'Tasks.Read', 'Mail.ReadBasic'];

export const loginRequest: PopupRequest = { scopes: graphScopes };

export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
