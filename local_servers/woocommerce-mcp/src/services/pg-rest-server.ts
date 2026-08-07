import http from 'node:http';
import { PgRestRouter } from './pg-rest-router.js';

// WooCommerce PG-backed REST facade.
//
// The WooCommerce MCP services (src/services/*.ts) speak to a real
// WooCommerce REST API at `${WORDPRESS_SITE_URL}/wp-json/wc/v3/...` via axios.
// In the evaluation image there is no WordPress/PHP backend; instead the WC
// data lives in Postgres (schema `wc`). This server bridges the two: it
// listens on the site URL port and routes REST calls to PgRestRouter, which
// translates them into SQL against the per-task Postgres instance.
//
// All WooCommerce tasks share the same `wc` schema, so a single generic
// wrapper works for every task; only the PG connection env vars differ.

const PORT = parseInt(process.env.WC_REST_PORT || '8081', 10);
const API_PREFIX = '/wp-json/wc/v3';

const router = new PgRestRouter();

function send(res: http.ServerResponse, status: number, body: any) {
  const payload = JSON.stringify(body === undefined ? null : body);
  res.writeHead(status, {
    'Content-Type': 'application/json',
    'Content-Length': Buffer.byteLength(payload),
  });
  res.end(payload);
}

function readBody(req: http.IncomingMessage): Promise<any> {
  return new Promise((resolve) => {
    const chunks: Buffer[] = [];
    req.on('data', (c) => chunks.push(c));
    req.on('end', () => {
      if (chunks.length === 0) return resolve({});
      try {
        resolve(JSON.parse(Buffer.concat(chunks).toString('utf8')));
      } catch {
        resolve({});
      }
    });
    req.on('error', () => resolve({}));
  });
}

const server = http.createServer(async (req, res) => {
  try {
    const url = new URL(req.url || '/', `http://localhost:${PORT}`);
    let path = url.pathname;

    // Health check
    if (path === '/health' || path === '/healthz') {
      return send(res, 200, { ok: true, backend: 'pg' });
    }

    // Only handle the WooCommerce REST namespace; anything else -> 404.
    if (!path.startsWith(API_PREFIX)) {
      return send(res, 404, { code: 'rest_no_route', message: 'No route was found matching the URL and request method.' });
    }
    // Strip the prefix -> e.g. "/orders/123" -> "orders/123"
    path = path.slice(API_PREFIX.length).replace(/^\/+/, '');

    // Collect query params (axios passes them in the query string).
    const params: Record<string, any> = {};
    url.searchParams.forEach((value, key) => {
      params[key] = value;
    });

    const method = (req.method || 'GET').toUpperCase();
    let result: any;

    if (method === 'GET') {
      result = await router.get(path, { params });
    } else if (method === 'POST') {
      const body = await readBody(req);
      result = await router.post(path, body, { params });
    } else if (method === 'PUT' || method === 'PATCH') {
      const body = await readBody(req);
      result = await router.put(path, body, { params });
    } else if (method === 'DELETE') {
      const body = await readBody(req);
      result = await router.delete(path, { params, data: body });
    } else {
      return send(res, 405, { code: 'rest_method_not_allowed', message: `Method ${method} not allowed.` });
    }

    // PgRestRouter returns an axios-like { data, status } envelope.
    const status = result?.status || 200;
    send(res, status, result?.data !== undefined ? result.data : result);
  } catch (err: any) {
    console.error('[pg-rest-server] error:', err?.message || err);
    send(res, 500, { code: 'woocommerce_rest_error', message: err?.message || 'Internal error' });
  }
});

server.listen(PORT, '127.0.0.1', () => {
  console.error(`[pg-rest-server] WooCommerce PG REST facade listening on 127.0.0.1:${PORT}`);
});
