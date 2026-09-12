import type { Configuration, PopupRequest } from '@azure/msal-browser';
import { LogLevel } from '@azure/msal-browser';

function required(name: string, value: string | undefined): string {
  if (!value || value.startsWith('00000000-')) {
    throw new Error(
      `${name} is not set. Run start.cmd, or edit config.json beside the page, ` +
        'with the values from the Entra ID app registration (see README).',
    );
  }
  return value;
}

export interface HubConfig {
  clientId?: string;
  tenantId?: string;
  redirectUri?: string;
}

/**
 * Configuration read at startup, not compiled in.
 *
 * Vite replaces import.meta.env.VITE_* at build time, so a bundle built without
 * the IDs is permanently stuck without them. config.json is fetched when the app
 * starts instead, which lets one build serve any tenant: host the bundle, edit
 * one text file beside it, done. The .env values remain the fallback, so
 * `npm run dev` keeps working exactly as before.
 */
let runtime: HubConfig = {};

const fromEnv = (key: string): string | undefined => {
  const value = (import.meta.env as Record<string, string | undefined>)[key];
  return value && !value.startsWith('00000000-') ? value : undefined;
};

const clean = (value: string | undefined): string | undefined =>
  value && value.trim() && !value.trim().startsWith('00000000-') ? value.trim() : undefined;

/** Fetch config.json beside the page. A missing or unreadable file is not an error. */
export async function loadRuntimeConfig(): Promise<void> {
  if (typeof fetch !== 'function') return;
  try {
    const response = await fetch('./config.json', { cache: 'no-store' });
    if (!response.ok) return;
    const parsed = (await response.json()) as HubConfig;
    runtime = {
      clientId: clean(parsed.clientId),
      tenantId: clean(parsed.tenantId),
      redirectUri: clean(parsed.redirectUri),
    };
  } catch {
    // No config.json, or it is not JSON. The .env values still apply.
  }
}

export const getClientId = (): string | undefined =>
  runtime.clientId ?? fromEnv('VITE_CLIENT_ID');

export const getTenantId = (): string | undefined =>
  runtime.tenantId ?? fromEnv('VITE_TENANT_ID');

/** Resolved lazily: `window` does not exist when the module loads under a test runner. */
export function resolveRedirectUri(): string {
  const configured = runtime.redirectUri ?? fromEnv('VITE_REDIRECT_URI');
  if (configured) return configured;
  return typeof window === 'undefined' ? 'http://localhost:5173' : window.location.origin;
}

/** True once a real client and tenant ID have been found, from either source. */
export function isConfigured(): boolean {
  try {
    required('VITE_CLIENT_ID', getClientId());
    required('VITE_TENANT_ID', getTenantId());
    return true;
  } catch {
    return false;
  }
}

export function buildMsalConfig(): Configuration {
  return {
    auth: {
      clientId: required('VITE_CLIENT_ID', getClientId()),
      authority: `https://login.microsoftonline.com/${required('VITE_TENANT_ID', getTenantId())}`,
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
 * Delegated scopes.
 *
 * Tasks.ReadWrite covers both the read and the two writes the Hub makes, complete
 * and reschedule; Graph offers no narrower permission for updating a task. Mail
 * stays read-only: the Hub never writes to the message resource. Remove
 * Mail.ReadBasic here and from the app registration to drop sender names and run
 * on Microsoft To Do data alone.
 */
export const graphScopes = [
  'User.Read',
  'Tasks.ReadWrite',
  'Mail.ReadBasic',
  'Calendars.ReadBasic',
];

export const loginRequest: PopupRequest = { scopes: graphScopes };

export const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
