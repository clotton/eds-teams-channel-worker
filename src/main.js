// Cloudflare Worker backend for Microsoft Teams with rate limiting handling

export default {
  async fetch(request, env, ctx) {
    const { searchParams } = new URL(request.url);
    const token = await getGraphToken(env);
    const teamId = searchParams.get("teamId");

    if (!teamId) {
      return new Response("Missing teamId", { status: 400 });
    }

    const team = await safeFetchGraph(`teams/${teamId}`, token);
    const channels = await safeFetchGraph(`teams/${teamId}/channels`, token);

    const stats = await limitConcurrency(
      channels.value.map(channel => () => getMessageStats(teamId, channel.id, token)),
      5
    );

    return Response.json({
      team,
      channels: channels.value,
      stats
    });
  }
};

async function getGraphToken(env) {
  const res = await fetch(env.AUTH_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: env.CLIENT_ID,
      client_secret: env.CLIENT_SECRET,
      scope: "https://graph.microsoft.com/.default",
      grant_type: "client_credentials"
    })
  });

  const json = await res.json();
  return json.access_token;
}

async function safeFetchGraph(endpoint, token, retries = 3, delay = 500) {
  for (let attempt = 0; attempt <= retries; attempt++) {
    const res = await fetch(`https://graph.microsoft.com/v1.0/${endpoint}`, {
      headers: { Authorization: `Bearer ${token}` }
    });

    if (res.status === 429) {
      const retryAfter = parseInt(res.headers.get("Retry-After") || "1", 10) * 1000;
      await new Promise(r => setTimeout(r, retryAfter || delay));
    } else if (res.ok) {
      return res.json();
    } else if (attempt === retries) {
      console.error(`Failed after ${retries + 1} attempts: ${endpoint}`);
      return null;
    } else {
      await new Promise(r => setTimeout(r, delay));
    }
  }
}

function limitConcurrency(tasks, limit) {
  const results = [];
  let i = 0;

  return new Promise((resolve) => {
    const run = () => {
      if (i >= tasks.length) return;
      const currentIndex = i++;
      tasks[currentIndex]()
        .then((res) => results[currentIndex] = res)
        .finally(() => {
          if (results.length === tasks.length) resolve(results);
          else run();
        });
    };

    for (let j = 0; j < limit; j++) run();
  });
}

async function getMessageStats(teamId, channelId, token) {
  const messages = await safeFetchGraph(`teams/${teamId}/channels/${channelId}/messages`, token);
  if (!messages || !messages.value) return { channelId, messageCount: 0 };

  const users = new Set();
  for (const msg of messages.value) {
    if (msg.from?.user?.id) users.add(msg.from.user.id);
  }

  return {
    channelId,
    messageCount: messages.value.length,
    uniqueSenders: users.size
  };
}
