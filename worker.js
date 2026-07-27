// Shared-mailbox reply auto-tagger for Microsoft 365
// Runs as a Cloudflare Worker. See README.md for full setup instructions.
//
// What it does:
//   - Listens for Microsoft Graph change notifications on a shared mailbox's Sent Items.
//   - When a reply goes out "on behalf of" the shared mailbox, it works out who the real
//     sender was (Graph exposes this natively for on-behalf sends) and tags that message
//     with a category matching their name.
//   - It also finds the original message the reply is threaded to (same conversationId)
//     and tags that too, so the whole thread shows who handled it at a glance.
//   - A daily Cron Trigger keeps the Graph subscription alive (they expire after 7 days
//     max and Microsoft does not auto-renew them).
//
// Required Worker secrets (Settings -> Variables and Secrets, mark as "Secret"):
//   GRAPH_TENANT_ID     - Entra "Directory (tenant) ID"
//   GRAPH_CLIENT_ID     - Entra app "Application (client) ID"
//   GRAPH_CLIENT_SECRET - the client secret VALUE (not the Secret ID)
//   GRAPH_CLIENT_STATE  - a random string you generate yourself, used to verify
//                         incoming webhook calls actually came from your subscription
//
// Required Worker plain text variables:
//   MAILBOX          - the shared mailbox address, e.g. office@example.com
//   NOTIFICATION_URL - the public URL this Worker is reachable at, e.g.
//                       https://automation.example.com/notifications
//
// Edit CATEGORY_MAP below to match your team and the category names you've already
// created in Outlook for the shared mailbox.

export default {
  async fetch(request, env, ctx) {
    const url = new URL(request.url);

    // Microsoft Graph validates a new subscription by calling your URL with
    // ?validationToken=... and expects it echoed back as plain text within seconds.
    const validationToken = url.searchParams.get('validationToken');
    if (validationToken !== null) {
      return new Response(validationToken, { status: 200, headers: { 'Content-Type': 'text/plain' } });
    }

    if (url.pathname === '/notifications' && request.method === 'POST') {
      let body;
      try { body = await request.json(); } catch (e) { body = null; }
      if (body && Array.isArray(body.value)) {
        // Graph expects a fast 202 ack; do the actual work in the background.
        ctx.waitUntil(processNotifications(body.value, env));
      }
      return new Response(null, { status: 202 });
    }

    // Manual trigger to create/renew the Graph subscription, e.g. right after first
    // deploying, instead of waiting for the daily Cron Trigger to fire.
    // Visit: https://your-worker-url/status?key=<your GRAPH_CLIENT_STATE value>
    if (url.pathname === '/status' && request.method === 'GET') {
      if (url.searchParams.get('key') !== env.GRAPH_CLIENT_STATE) {
        return new Response('Unauthorized', { status: 401 });
      }
      try {
        const result = await ensureSubscription(env);
        return new Response(JSON.stringify(result, null, 2), { status: 200, headers: { 'Content-Type': 'application/json' } });
      } catch (err) {
        return new Response('Error: ' + err.message, { status: 500 });
      }
    }

    return new Response('Not found', { status: 404 });
  },

  async scheduled(event, env, ctx) {
    ctx.waitUntil(ensureSubscription(env));
  },
};

// Map the sender's display name (as Microsoft Graph reports it, lowercased) to the
// exact category name you've already created in Outlook for the shared mailbox.
const CATEGORY_MAP = {
  'alice example': '✅ Alice',
  'bob example': '✅ Bob',
  'carol example': '✅ Carol',
};

async function getAccessToken(env) {
  const resp = await fetch('https://login.microsoftonline.com/' + env.GRAPH_TENANT_ID + '/oauth2/v2.0/token', {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: env.GRAPH_CLIENT_ID,
      client_secret: env.GRAPH_CLIENT_SECRET,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    }),
  });
  if (!resp.ok) {
    throw new Error('Token request failed: ' + resp.status + ' ' + await resp.text());
  }
  const data = await resp.json();
  return data.access_token;
}

async function graphFetch(path, token, options) {
  options = options || {};
  return fetch('https://graph.microsoft.com/v1.0' + path, {
    ...options,
    headers: {
      Authorization: 'Bearer ' + token,
      'Content-Type': 'application/json',
      ...(options.headers || {}),
    },
  });
}

async function processNotifications(notifications, env) {
  let token;
  try {
    token = await getAccessToken(env);
  } catch (err) {
    console.error('Failed to get access token', err.message);
    return;
  }
  for (const n of notifications) {
    try {
      // Reject anything that doesn't carry the shared secret we set when creating
      // the subscription - stops spoofed calls to this endpoint from doing anything.
      if (env.GRAPH_CLIENT_STATE && n.clientState !== env.GRAPH_CLIENT_STATE) {
        console.error('clientState mismatch, ignoring notification');
        continue;
      }
      const messageId = n.resourceData && n.resourceData.id;
      if (!messageId) continue;
      await categorizeMessageAndThread(messageId, token, env);
    } catch (err) {
      console.error('Error processing notification', err.message);
    }
  }
}

async function categorizeMessageAndThread(messageId, token, env) {
  const mailbox = env.MAILBOX;
  const msgResp = await graphFetch(
    '/users/' + encodeURIComponent(mailbox) + '/messages/' + messageId + '?$select=id,sender,from,conversationId,categories',
    token
  );
  if (!msgResp.ok) {
    console.error('Failed to fetch message', messageId, msgResp.status, await msgResp.text());
    return;
  }
  const msg = await msgResp.json();

  // For "send on behalf of" messages, Graph's `sender` is the real person and `from`
  // is the shared mailbox. For a normal send (or Send As), the two are identical -
  // that's exactly the case we want to skip, since there's no real sender to tag.
  const senderEmail = msg.sender && msg.sender.emailAddress && msg.sender.emailAddress.address
    ? msg.sender.emailAddress.address.toLowerCase() : null;
  const fromEmail = msg.from && msg.from.emailAddress && msg.from.emailAddress.address
    ? msg.from.emailAddress.address.toLowerCase() : null;
  const senderName = msg.sender && msg.sender.emailAddress && msg.sender.emailAddress.name
    ? msg.sender.emailAddress.name.toLowerCase() : null;

  console.log('Processing message', messageId, 'sender:', senderEmail, 'from:', fromEmail, 'senderName:', senderName, 'conversationId:', msg.conversationId);

  if (!senderEmail || !fromEmail || senderEmail === fromEmail) {
    console.log('Not a send-on-behalf message, skipping');
    return;
  }

  const category = senderName ? CATEGORY_MAP[senderName] : null;
  if (!category) {
    console.error('No category mapped for sender name:', senderName);
    return;
  }

  await addCategory(mailbox, msg.id, msg.categories || [], category, token);
  console.log('Tagged sent item', msg.id, 'with', category);

  // Find every other message in the same conversation (e.g. the original message
  // sitting in the Inbox) and tag those too, so the whole thread shows who replied.
  // Note: Graph does not support filtering on `id`, so we filter by conversationId
  // only and exclude the current message client-side afterwards.
  if (msg.conversationId) {
    const filterQuery = "conversationId eq '" + msg.conversationId + "'";
    const convResp = await graphFetch(
      '/users/' + encodeURIComponent(mailbox) + '/messages?$filter=' + encodeURIComponent(filterQuery) + '&$select=id,categories',
      token
    );
    if (!convResp.ok) {
      console.error('Conversation lookup failed', convResp.status, await convResp.text());
      return;
    }
    const convData = await convResp.json();
    const others = (convData.value || []).filter(item => item.id !== msg.id);
    console.log('Conversation matches found (excluding self):', others.length);
    for (const item of others) {
      await addCategory(mailbox, item.id, item.categories || [], category, token);
      console.log('Tagged conversation item', item.id, 'with', category);
    }
  } else {
    console.log('No conversationId on message, skipping thread tagging');
  }
}

async function addCategory(mailbox, messageId, existingCategories, category, token) {
  if (existingCategories.includes(category)) return;
  const updated = existingCategories.concat([category]);
  const resp = await graphFetch('/users/' + encodeURIComponent(mailbox) + '/messages/' + messageId, token, {
    method: 'PATCH',
    body: JSON.stringify({ categories: updated }),
  });
  if (!resp.ok) {
    console.error('Failed to patch categories on', messageId, resp.status, await resp.text());
  }
}

async function ensureSubscription(env) {
  const token = await getAccessToken(env);
  const listResp = await graphFetch('/subscriptions', token);
  const list = listResp.ok ? await listResp.json() : { value: [] };
  const existing = (list.value || []).find(s => s.notificationUrl === env.NOTIFICATION_URL);

  // Outlook message subscriptions max out at 10,080 minutes (7 days) with no auto-renew.
  // We ask for ~6 days and renew daily via the scheduled() handler, well inside the window.
  const expiry = new Date(Date.now() + 6 * 24 * 60 * 60 * 1000).toISOString();

  if (existing) {
    const patchResp = await graphFetch('/subscriptions/' + existing.id, token, {
      method: 'PATCH',
      body: JSON.stringify({ expirationDateTime: expiry }),
    });
    return { action: 'renewed', ok: patchResp.ok, status: patchResp.status, id: existing.id };
  } else {
    const createResp = await graphFetch('/subscriptions', token, {
      method: 'POST',
      body: JSON.stringify({
        changeType: 'created',
        notificationUrl: env.NOTIFICATION_URL,
        resource: "/users/" + encodeURIComponent(env.MAILBOX) + "/mailFolders('SentItems')/messages",
        expirationDateTime: expiry,
        clientState: env.GRAPH_CLIENT_STATE,
      }),
    });
    const body = await createResp.text();
    return { action: 'created', ok: createResp.ok, status: createResp.status, body };
  }
}
