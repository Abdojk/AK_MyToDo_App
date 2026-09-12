import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { PublicClientApplication } from '@azure/msal-browser';
import { App } from './App';
import { buildMsalConfig, isConfigured } from './auth/msalConfig';
import { SignInPanel } from './components/SignInPanel';
import './styles.css';

const container = document.getElementById('root');
if (!container) throw new Error('Root element missing from index.html.');

const root = createRoot(container);

async function start(): Promise<void> {
  if (!isConfigured()) {
    root.render(
      <StrictMode>
        <main className="shell shell-centred">
          <SignInPanel onSignIn={() => {}} busy={false} error={null} configured={false} />
        </main>
      </StrictMode>,
    );
    return;
  }

  const msal = new PublicClientApplication(buildMsalConfig());
  await msal.initialize();
  await msal.handleRedirectPromise();

  root.render(
    <StrictMode>
      <App msal={msal} />
    </StrictMode>,
  );
}

void start();
