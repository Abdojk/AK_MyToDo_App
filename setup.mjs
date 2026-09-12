#!/usr/bin/env node
/**
 * One-step launcher for the AK MyToDo Hub.
 *
 * Installs dependencies if they are missing, asks for the two IDs from the Entra
 * registration if .env does not yet hold them, starts the development server, and
 * opens the browser. Running it a second time skips straight to the server.
 *
 * Usage: node setup.mjs     (or double-click start.cmd on Windows)
 */

import { spawn, spawnSync } from 'node:child_process';
import { createInterface } from 'node:readline/promises';
import { existsSync, readFileSync, writeFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';
import { stdin, stdout } from 'node:process';

const root = dirname(fileURLToPath(import.meta.url));
const envPath = join(root, '.env');
const PORT = 5173;
const URL = `http://localhost:${PORT}`;

const GUID = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;
const DOMAIN = /^[a-z0-9.-]+\.[a-z]{2,}$/i;
const PLACEHOLDER = '00000000-0000-0000-0000-000000000000';

const say = (line = '') => stdout.write(`${line}\n`);

function fail(message) {
  say('');
  say(`  ${message}`);
  say('');
  process.exit(1);
}

/** Node 20 or later, because the app's toolchain needs it. */
function checkNode() {
  const major = Number(process.versions.node.split('.')[0]);
  if (major < 20) {
    fail(
      `This needs Node 20 or later. You are running ${process.version}.\n` +
        '  Install the current release from https://nodejs.org, then run this again.',
    );
  }
}

function run(command, args) {
  const result = spawnSync(command, args, {
    cwd: root,
    stdio: 'inherit',
    shell: process.platform === 'win32',
  });
  if (result.status !== 0) fail(`"${command} ${args.join(' ')}" failed.`);
}

function installIfNeeded() {
  if (existsSync(join(root, 'node_modules'))) return;
  say('Installing dependencies. This takes a minute the first time.');
  say('');
  run('npm', ['install']);
  say('');
}

/** True once .env carries something other than the placeholder GUIDs. */
function envIsConfigured() {
  if (!existsSync(envPath)) return false;
  const text = readFileSync(envPath, 'utf8');
  const value = (key) => text.match(new RegExp(`^${key}=(.*)$`, 'm'))?.[1]?.trim() ?? '';
  const client = value('VITE_CLIENT_ID');
  const tenant = value('VITE_TENANT_ID');
  return Boolean(client) && client !== PLACEHOLDER && Boolean(tenant) && tenant !== PLACEHOLDER;
}

/** A value passed as --client <x>, or in the environment. */
function suppliedValue(flag, envKey) {
  const index = process.argv.indexOf(`--${flag}`);
  if (index !== -1 && process.argv[index + 1]) return process.argv[index + 1].trim();
  return process.env[envKey]?.trim() ?? '';
}

const isClientId = (v) => GUID.test(v);
const isTenantId = (v) => GUID.test(v) || DOMAIN.test(v);

/**
 * Ask once, then keep asking until the value looks right.
 *
 * A closed stdin resolves to null rather than hanging, which is what happens when
 * the script is piped or run by a task scheduler instead of typed into.
 */
async function askUntilValid(rl, prompt, isValid, complaint) {
  for (;;) {
    const answer = await rl.question(prompt);
    if (answer === null || answer === undefined) return null;
    const value = answer.trim();
    if (isValid(value)) return value;
    say(`  ${complaint}`);
  }
}

async function askForIds() {
  let clientId = suppliedValue('client', 'VITE_CLIENT_ID');
  let tenantId = suppliedValue('tenant', 'VITE_TENANT_ID');

  if (!isClientId(clientId) || !isTenantId(tenantId)) {
    if (!stdin.isTTY) {
      fail(
        'The two IDs from the Entra registration are missing, and there is no\n' +
          '  terminal to ask on. Either run this from a command prompt, or pass them:\n' +
          '    node setup.mjs --client <application-id> --tenant <directory-id>',
      );
    }

    say('');
    say('  Two values are needed, both from the Entra ID app registration.');
    say("  Find them on the registration's Overview page. The README covers");
    say('  creating the registration, under "Step 0".');
    say('');

    const rl = createInterface({ input: stdin, output: stdout });
    // A Ctrl+C or a closed input leaves nothing to write, so stop rather than
    // writing a half-finished .env the app would then reject.
    rl.on('close', () => {});

    if (!isClientId(clientId)) {
      clientId = await askUntilValid(
        rl,
        '  Application (client) ID: ',
        isClientId,
        'That is not a GUID. It looks like 11111111-2222-3333-4444-555555555555.',
      );
    }
    if (clientId && !isTenantId(tenantId)) {
      tenantId = await askUntilValid(
        rl,
        '  Directory (tenant) ID:   ',
        isTenantId,
        'That is neither a GUID nor a domain such as info-sys.com.',
      );
    }
    rl.close();

    if (!clientId || !tenantId) fail('Cancelled. Nothing was written.');
  }

  writeFileSync(
    envPath,
    [
      '# Written by setup.mjs from the Entra ID app registration.',
      `VITE_CLIENT_ID=${clientId}`,
      `VITE_TENANT_ID=${tenantId}`,
      `VITE_REDIRECT_URI=${URL}`,
      '',
    ].join('\n'),
    'utf8',
  );

  say('');
  say(`  Saved to ${envPath}`);
}

function openBrowser() {
  const [command, args] =
    process.platform === 'win32'
      ? ['cmd', ['/c', 'start', '', URL]]
      : process.platform === 'darwin'
        ? ['open', [URL]]
        : ['xdg-open', [URL]];
  // spawn reports a missing command through an async 'error' event, not a throw.
  // Without this listener that event is unhandled and takes the server down with
  // it, which reads as the Hub crashing on startup. Failing to open a browser is
  // never worth that: the URL is on screen either way.
  const child = spawn(command, args, { detached: true, stdio: 'ignore' });
  child.on('error', () => say(`  Open ${URL} yourself; this machine has no browser command.`));
  child.unref();
}

/**
 * Whether something already answers on the port.
 *
 * The redirect URI in the Entra registration names port 5173, so the server
 * cannot simply move to another one: sign-in would fail. A clear message beats
 * Vite's raw "Port 5173 is already in use".
 */
async function portIsTaken() {
  try {
    await fetch(URL, { signal: AbortSignal.timeout(1500) });
    return true;
  } catch {
    return false;
  }
}

async function startServer() {
  if (await portIsTaken()) {
    fail(
      `Something is already using port ${PORT}.\n` +
        `  It may be the Hub itself: try ${URL} before anything else.\n` +
        '  Otherwise close whatever holds the port, then run this again.\n' +
        '  The port is fixed because the Entra redirect URI names it.',
    );
  }

  say('');
  say(`  Starting the Hub on ${URL}`);
  say('  Leave this window open while you use it. Press Ctrl+C to stop.');
  say('');

  const server = spawn('npm', ['run', 'dev'], {
    cwd: root,
    stdio: 'inherit',
    shell: process.platform === 'win32',
  });

  // Wait for the server to answer before opening the browser, so the first load
  // is not a connection error.
  for (let attempt = 0; attempt < 40; attempt += 1) {
    await new Promise((resolve) => setTimeout(resolve, 250));
    try {
      const response = await fetch(URL);
      if (response.ok) {
        openBrowser();
        break;
      }
    } catch {
      // Not up yet.
    }
  }

  server.on('exit', (code) => process.exit(code ?? 0));
}

async function main() {
  say('');
  say('  AK MyToDo Hub');
  say('  =============');
  checkNode();
  installIfNeeded();
  if (!envIsConfigured()) await askForIds();
  await startServer();
}

await main();
