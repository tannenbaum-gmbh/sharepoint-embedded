const graphRoot = 'https://graph.microsoft.com/v1.0';
const guid = /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

export function loadConfig(env = process.env) {
  for (const name of ['TENANT_ID', 'CLIENT_ID', 'CONTAINER_TYPE_ID']) {
    if (!guid.test(env[name] || '')) throw new Error(`${name} must be a GUID in .env.`);
  }
  if (!env.CLIENT_SECRET) throw new Error('Set CLIENT_SECRET in .env.');
  return {
    tenantId: env.TENANT_ID,
    clientId: env.CLIENT_ID,
    clientSecret: env.CLIENT_SECRET,
    containerTypeId: env.CONTAINER_TYPE_ID
  };
}

export function createGraphClient(config, fetchImpl = fetch) {
  let token;
  let expiresAt = 0;

  async function request(path, options = {}) {
    const url = new URL(path, `${graphRoot}/`);
    // Graph pagination must never send the app token to another host.
    if (url.origin !== 'https://graph.microsoft.com' || !url.pathname.startsWith('/v1.0/')) {
      throw new Error('Invalid Graph URL.');
    }
    if (!token || Date.now() >= expiresAt) {
      const response = await fetchImpl(
        `https://login.microsoftonline.com/${config.tenantId}/oauth2/v2.0/token`,
        {
          method: 'POST',
          body: new URLSearchParams({
            client_id: config.clientId,
            client_secret: config.clientSecret,
            grant_type: 'client_credentials',
            scope: 'https://graph.microsoft.com/.default'
          }),
          signal: AbortSignal.timeout(30000),
          redirect: 'error'
        }
      );
      if (!response.ok) throw new Error(`Entra authentication failed (${response.status}). Check credentials and consent.`);
      const data = await response.json();
      if (!data.access_token || !Number.isFinite(Number(data.expires_in))) {
        throw new Error('Entra returned an invalid token response.');
      }
      token = data.access_token;
      expiresAt = Date.now() + (Number(data.expires_in) - 60) * 1000;
    }
    const response = await fetchImpl(url, {
      ...options,
      headers: { ...options.headers, Authorization: 'Bearer ' + token },
      signal: AbortSignal.timeout(30000),
      redirect: 'error'
    });
    if (!response.ok) {
      if (response.status === 401) expiresAt = 0;
      throw new Error(`Graph request failed (${response.status}). Check permissions, registration, and trial expiry.`);
    }
    return response.status === 204 ? null : response.json();
  }

  async function list(path) {
    const value = [];
    for (let page = 0; path; page++) {
      if (page >= 100) throw new Error('Too many Graph pages for this test app.');
      const data = await request(path);
      value.push(...data.value);
      path = data['@odata.nextLink'];
    }
    return { value };
  }

  return { request, list };
}
