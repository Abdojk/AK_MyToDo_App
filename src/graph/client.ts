import { GRAPH_BASE } from '../auth/msalConfig';

export class GraphError extends Error {
  constructor(
    message: string,
    readonly status: number,
    readonly code: string | null,
    readonly url: string,
  ) {
    super(message);
    this.name = 'GraphError';
  }
}

interface GraphErrorBody {
  error?: { code?: string; message?: string };
}

interface BatchRequest {
  id: string;
  method: 'GET';
  url: string;
}

interface BatchResponse<T> {
  id: string;
  status: number;
  body?: T & GraphErrorBody;
}

const MAX_RETRIES = 3;
const BATCH_SIZE = 20;

function delay(ms: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

/**
 * Microsoft Graph client.
 *
 * Reads issue GET, and the one POST is to /$batch carrying GET requests only. The
 * single write verb is PATCH, which reaches only the To Do task endpoints in
 * graph/write.ts. Nothing here touches the message resource with a write.
 */
export class GraphClient {
  constructor(private readonly getToken: () => Promise<string>) {}

  private async request(
    url: string,
    init: RequestInit = {},
    attempt = 0,
  ): Promise<Response> {
    const token = await this.getToken();
    const response = await fetch(url, {
      ...init,
      headers: {
        ...init.headers,
        Authorization: `Bearer ${token}`,
        Accept: 'application/json',
      },
    });

    // Graph throttles per mailbox. Honour Retry-After, then back off.
    if ((response.status === 429 || response.status >= 500) && attempt < MAX_RETRIES) {
      const retryAfter = Number(response.headers.get('Retry-After'));
      const wait = Number.isFinite(retryAfter) && retryAfter > 0
        ? retryAfter * 1000
        : 2 ** attempt * 1000;
      await delay(wait);
      return this.request(url, init, attempt + 1);
    }

    return response;
  }

  private static absolute(pathOrUrl: string): string {
    return pathOrUrl.startsWith('http') ? pathOrUrl : `${GRAPH_BASE}${pathOrUrl}`;
  }

  private static async toError(response: Response, url: string): Promise<GraphError> {
    let code: string | null = null;
    let message = `Microsoft Graph returned ${response.status}.`;
    try {
      const body = (await response.json()) as GraphErrorBody;
      code = body.error?.code ?? null;
      if (body.error?.message) message = body.error.message;
    } catch {
      // A non-JSON error body leaves the status line as the message.
    }
    return new GraphError(message, response.status, code, url);
  }

  /** Single GET. */
  async get<T>(path: string): Promise<T> {
    const url = GraphClient.absolute(path);
    const response = await this.request(url);
    if (!response.ok) throw await GraphClient.toError(response, url);
    return (await response.json()) as T;
  }

  /** Single PATCH. The only write the client makes. */
  async patch<T>(path: string, body: unknown): Promise<T> {
    const url = GraphClient.absolute(path);
    const response = await this.request(url, {
      method: 'PATCH',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body),
    });
    if (!response.ok) throw await GraphClient.toError(response, url);
    return (await response.json()) as T;
  }

  /** GET a collection, following @odata.nextLink until the server stops paging. */
  async getAll<T>(path: string): Promise<T[]> {
    const items: T[] = [];
    let next: string | undefined = GraphClient.absolute(path);
    while (next) {
      const page: { value?: T[]; '@odata.nextLink'?: string } = await this.get(next);
      if (page.value) items.push(...page.value);
      next = page['@odata.nextLink'];
    }
    return items;
  }

  /**
   * Run many GETs through /$batch, twenty per request.
   *
   * Returns a map from the caller's request id to the response body. A failed
   * entry is left out rather than failing the whole read, so one unreadable
   * task never blanks the timeline.
   */
  async batchGet<T>(requests: BatchRequest[]): Promise<Map<string, T>> {
    const results = new Map<string, T>();

    for (let i = 0; i < requests.length; i += BATCH_SIZE) {
      const chunk = requests.slice(i, i + BATCH_SIZE);
      const url = `${GRAPH_BASE}/$batch`;
      const response = await this.request(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ requests: chunk }),
      });
      if (!response.ok) throw await GraphClient.toError(response, url);

      const payload = (await response.json()) as { responses?: BatchResponse<T>[] };
      for (const entry of payload.responses ?? []) {
        if (entry.status >= 200 && entry.status < 300 && entry.body) {
          results.set(entry.id, entry.body);
        } else {
          console.warn(
            `[graph] batch entry ${entry.id} returned ${entry.status}`,
            entry.body?.error?.code ?? '',
          );
        }
      }
    }

    return results;
  }
}
