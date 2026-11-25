import test from 'node:test';
import assert from 'node:assert/strict';
import { request as httpRequest } from 'node:http';
import { createGraphClient, loadConfig } from '../app/graph.js';
import { createApp } from '../app/server.js';

const id = '11111111-1111-1111-1111-111111111111';
const config = { tenantId: id, clientId: id, clientSecret: 'local-test-placeholder', containerTypeId: id };

test('configuration requires GUIDs and a secret', () => {
  assert.throws(() => loadConfig({}), /TENANT_ID/);
  assert.throws(() => loadConfig({ TENANT_ID: id, CLIENT_ID: id, CONTAINER_TYPE_ID: id }), /CLIENT_SECRET/);
  assert.deepEqual(loadConfig({
    TENANT_ID: id, CLIENT_ID: id, CONTAINER_TYPE_ID: id, CLIENT_SECRET: config.clientSecret
  }), config);
});

test('client credentials stay server-side, tokens are cached, and pagination works', async () => {
  const calls = [];
  const client = createGraphClient(config, async (url, options) => {
    calls.push({ url: String(url), options });
    if (String(url).includes('/token')) {
      assert.equal(options.body.get('grant_type'), 'client_credentials');
      assert.equal(options.body.get('client_secret'), config.clientSecret);
      assert.equal(options.body.get('scope'), 'https://graph.microsoft.com/.default');
      return Response.json({ access_token: 'mock-access-token', expires_in: 3600 });
    }
    assert.equal(options.headers.Authorization, 'Bearer ' + 'mock-access-token');
    assert.equal(options.redirect, 'error');
    return calls.length === 2
      ? Response.json({ value: [{ id: 'first' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/next' })
      : Response.json({ value: [{ id: 'second' }] });
  });
  assert.deepEqual(await client.list('storage/fileStorage/containers'), {
    value: [{ id: 'first' }, { id: 'second' }]
  });
  assert.equal(calls.length, 3);
  await assert.rejects(client.request('https://example.com/steal'), /Invalid Graph URL/);
  assert.equal(calls.length, 3);
});

test('authentication errors never include the remote response or secret', async () => {
  const client = createGraphClient(config, async () => Response.json({ error: config.clientSecret }, { status: 401 }));
  await assert.rejects(client.request('storage/fileStorage/containers'), error => {
    assert.match(error.message, /authentication failed \(401\)/);
    assert.ok(!error.message.includes(config.clientSecret));
    return true;
  });
});

async function fixture(t) {
  const calls = [];
  const graph = {
    async list(path) {
      calls.push({ path });
      return { value: path.includes('children')
        ? [{ id: 'file', name: '<script>.txt', size: 5, '@microsoft.graph.downloadUrl': 'sensitive-url' }]
        : [{ id: 'container', displayName: 'Test' }] };
    },
    async request(path, options) {
      calls.push({ path, options });
      if (!options) return { containerTypeId: path.endsWith('/other') ? 'different' : id };
      return { id: 'container', name: 'test.txt', size: 5, '@microsoft.graph.downloadUrl': 'sensitive-url' };
    }
  };
  const server = createApp(config, graph);
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  t.after(() => new Promise(resolve => {
    server.close(resolve);
    server.closeAllConnections();
  }));
  const origin = `http://127.0.0.1:${server.address().port}`;
  return { calls, origin, request: (path, options) => fetch(origin + path, options) };
}

test('serves the UI and lists only the configured container type', async t => {
  const { request, calls } = await fixture(t);
  const page = await request('/');
  assert.equal(page.status, 200);
  assert.match(await page.text(), /SharePoint Embedded test app/);
  assert.match(page.headers.get('content-security-policy'), /frame-ancestors 'none'/);
  const response = await request('/api/containers');
  assert.equal(response.status, 200);
  assert.equal(calls[0].path, `storage/fileStorage/containers?$filter=${encodeURIComponent(`containerTypeId eq ${id}`)}`);
});

test('creates a container through the backend with the configured type', async t => {
  const { request, origin, calls } = await fixture(t);
  const response = await request('/api/containers', {
    method: 'POST', headers: { Origin: origin }, body: JSON.stringify({ displayName: ' Test ' })
  });
  assert.equal(response.status, 201);
  assert.deepEqual(JSON.parse(calls[0].options.body), { displayName: 'Test', containerTypeId: id });
});

test('rejects cross-origin mutations, DNS rebinding, malformed JSON, and invalid names', async t => {
  const { request, origin, calls } = await fixture(t);
  assert.equal((await request('/api/containers', { method: 'POST', headers: { Origin: 'https://example.com' } })).status, 403);
  const reboundStatus = await new Promise(resolve => {
    httpRequest(origin + '/api/containers', { headers: { Host: 'example.com' } }, response => {
      response.resume();
      resolve(response.statusCode);
    }).end();
  });
  assert.equal(reboundStatus, 403);
  assert.equal((await request('/api/containers', { headers: { 'Sec-Fetch-Site': 'cross-site' } })).status, 403);
  for (const body of ['bad json', '{"displayName": ""}', 'null']) {
    assert.equal((await request('/api/containers', { method: 'POST', headers: { Origin: origin }, body })).status, 400);
  }
  assert.equal(calls.length, 0);
});

test('lists and uploads files without exposing download URLs', async t => {
  const { request, origin, calls } = await fixture(t);
  const listing = await request('/api/files?containerId=container');
  assert.deepEqual(await listing.json(), { value: [{ id: 'file', name: '<script>.txt', size: 5 }] });
  const upload = await request('/api/files?containerId=container&name=test.txt', {
    method: 'PUT', headers: { Origin: origin }, body: 'hello'
  });
  assert.equal(upload.status, 200);
  assert.deepEqual(await upload.json(), { id: 'container', name: 'test.txt', size: 5 });
  assert.equal(calls.at(-1).path, 'drives/container/root:/test.txt:/content');
  assert.equal(calls.at(-1).options.body.toString(), 'hello');
});

test('rejects foreign containers, path injection, invalid filenames, and oversized uploads', async t => {
  const { request, origin, calls } = await fixture(t);
  assert.equal((await request('/api/files?containerId=other')).status, 403);
  assert.equal((await request('/api/files?containerId=..%2Felsewhere')).status, 400);
  for (const name of ['../file', '.', 'bad#file', 'file.']) {
    assert.equal((await request(`/api/files?containerId=container&name=${encodeURIComponent(name)}`, {
      method: 'PUT', headers: { Origin: origin }, body: 'hello'
    })).status, 400);
  }
  assert.equal((await request('/api/files?containerId=container&name=big.txt', {
    method: 'PUT', headers: { Origin: origin }, body: Buffer.alloc(1024 * 1024 + 1)
  })).status, 413);
  assert.ok(!calls.some(call => call.options));
});
