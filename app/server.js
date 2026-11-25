import { createServer } from 'node:http';
import { readFile } from 'node:fs/promises';
import { fileURLToPath } from 'node:url';
import { resolve } from 'node:path';
import { createGraphClient, loadConfig } from './graph.js';

const maxUpload = 1024 * 1024;
const assets = {
  '/': ['index.html', 'text/html; charset=utf-8'],
  '/app.js': ['app.js', 'text/javascript; charset=utf-8'],
  '/style.css': ['style.css', 'text/css; charset=utf-8']
};

function badRequest(message, status = 400) {
  return Object.assign(new Error(message), { status });
}

async function body(req, limit) {
  const chunks = [];
  let size = 0;
  for await (const chunk of req) {
    size += chunk.length;
    if (size > limit) throw badRequest('Request exceeds the size limit.', 413);
    chunks.push(chunk);
  }
  return Buffer.concat(chunks);
}

function segment(value) {
  if (!value || value.length > 512 || !/^[a-zA-Z0-9!_-]+$/.test(value)) {
    throw badRequest('Invalid container ID.');
  }
  return encodeURIComponent(value);
}

export function createApp(config, graph = createGraphClient(config)) {
  return createServer(async (req, res) => {
    res.setHeader('Cache-Control', 'no-store');
    res.setHeader('X-Content-Type-Options', 'nosniff');
    res.setHeader('Content-Security-Policy', "default-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'self'");
    try {
      // Only private localhost forwards are supported; reject DNS rebinding and cross-site calls.
      if (!/^(localhost|127\.0\.0\.1)(:\d+)?$/.test(req.headers.host || '') ||
          (req.headers['sec-fetch-site'] && !['same-origin', 'none'].includes(req.headers['sec-fetch-site']))) {
        throw badRequest('Use a private localhost port forward.', 403);
      }
      if (req.method !== 'GET' && req.headers.origin !== `http://${req.headers.host}`) {
        throw badRequest('A same-origin request is required.', 403);
      }
      const url = new URL(req.url, `http://${req.headers.host}`);
      const asset = assets[url.pathname];
      if (req.method === 'GET' && asset) {
        res.setHeader('Content-Type', asset[1]);
        res.end(await readFile(new URL(`./public/${asset[0]}`, import.meta.url)));
        return;
      }
      let data;
      let status = 200;
      if (url.pathname === '/api/containers' && req.method === 'GET') {
        data = await graph.list(`storage/fileStorage/containers?$filter=${encodeURIComponent(`containerTypeId eq ${config.containerTypeId}`)}`);
      } else if (url.pathname === '/api/containers' && req.method === 'POST') {
        let input;
        try {
          input = JSON.parse((await body(req, 4096)).toString());
        } catch (error) {
          if (error.status) throw error;
          throw badRequest('Invalid JSON.');
        }
        if (typeof input?.displayName !== 'string' || !input.displayName.trim() || input.displayName.length > 100) {
          throw badRequest('Container name must be 1–100 characters.');
        }
        data = await graph.request('storage/fileStorage/containers', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            displayName: input.displayName.trim(),
            containerTypeId: config.containerTypeId
          })
        });
        status = 201;
      } else if (url.pathname === '/api/files' && ['GET', 'PUT'].includes(req.method)) {
        const id = segment(url.searchParams.get('containerId'));
        const container = await graph.request(`storage/fileStorage/containers/${id}`);
        if (container.containerTypeId?.toLowerCase() !== config.containerTypeId.toLowerCase()) {
          throw badRequest('Container belongs to a different container type.', 403);
        }
        if (req.method === 'GET') {
          const listing = await graph.list(`drives/${id}/root/children`);
          data = { value: listing.value.map(({ id, name, size, folder }) => ({ id, name, size, folder })) };
        } else {
          const name = url.searchParams.get('name');
          if (!name || name.length > 200 || /["*:<>?\/\\|#%\x00-\x1f]/.test(name) ||
              /[. ]$/.test(name) || /^\s/.test(name) || name === '.' || name === '..') {
            throw badRequest('Use a simple filename of 1–200 characters.');
          }
          data = await graph.request(`drives/${id}/root:/${encodeURIComponent(name)}:/content`, {
            method: 'PUT',
            headers: { 'Content-Type': 'application/octet-stream' },
            body: await body(req, maxUpload)
          });
          // Do not expose Graph preauthenticated download URLs.
          data = { id: data.id, name: data.name, size: data.size };
        }
      } else {
        throw badRequest('Not found.', 404);
      }
      res.writeHead(status, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify(data));
    } catch (error) {
      res.writeHead(error.status || 502, { 'Content-Type': 'application/json' });
      res.end(JSON.stringify({
        error: error.status ? error.message : 'Graph or authentication request failed. Check credentials, consent, registration, billing, and connectivity.'
      }));
    }
  });
}

if (process.argv[1] && resolve(process.argv[1]) === fileURLToPath(import.meta.url)) {
  try {
    const config = loadConfig();
    const port = Number(process.env.PORT || 3000);
    if (!Number.isInteger(port) || port < 1 || port > 65535) throw new Error('PORT must be 1–65535.');
    createApp(config).listen(port, '127.0.0.1', () => {
      console.log(`SharePoint Embedded test app: http://localhost:${port}`);
    });
  } catch (error) {
    console.error(error.message);
    process.exitCode = 1;
  }
}
