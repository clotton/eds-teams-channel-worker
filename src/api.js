export async function logMemberAddition({ addedBy, addedUser, teamName, added }, env) {
  const webhookUrl = env.SLACK_WEBHOOK_URL; // Replace with your webhook
  const message = {
    text: `👤 *${addedBy}* added *${addedUser}* to team *${teamName}* — ${added
      ? '✅ Success' : '❌ Failed'}`,
  };
  await fetch(webhookUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(message),
  });
}

export async function logSearchAttempt({ searchBy, searchName, searchDescription }, env) {
  const webhookUrl = env.SLACK_WEBHOOK_URL; // Replace with your webhook
  const message = {
    text: `👤 *${searchBy}* searched Teams for name: *${searchName}* and description: *${searchDescription}*`,
  };
  await fetch(webhookUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(message),
  });
}
export async function logMemberRemoval({ removedBy, removedUser, teamName, removed }, env) {
  const webhookUrl = env.SLACK_WEBHOOK_URL; // Replace with your webhook
  const message = {
    text: `👤 *${removedBy}* removed *${removedUser}* to team *${teamName}* — ${removed
      ? '✅ Success' : '❌ Failed'}`,
  };
  await fetch(webhookUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(message),
  });
}


const getUser = async (email, bearer) => {
  // prevent getting other users
  if (!email ||
      (
          !email.endsWith('@adobe.com') &&
          !email.toLowerCase().endsWith('@AdobeEnterpriseSupportAEM.onmicrosoft.com'.toLowerCase()))
  ) return null;

  const params = new URLSearchParams({
    '$filter': `endsWith(mail,'${email}')`,
    '$select': 'id,mail,displayName',
    '$count': 'true'
  });

  const headers = {
    ConsistencyLevel: 'eventual',
    Authorization: `Bearer ${bearer}`,
  };

  const url = `https://graph.microsoft.com/v1.0/users?${params}`

  const res = await fetch(url, {
    method: 'GET',
    headers,
  });

  const json = await res.json();
  if (json.value && json.value.length > 0) {
    return json.value[0];
  }

  return null;
};

const getUserTeams = async (data, env) => {
  const user = await getUser(data.id, data.bearer);
  if (!user) return null;

  const headers = {
    ConsistencyLevel: 'eventual',
    Authorization: `Bearer ${data.bearer}`,
  };

  const url = `https://graph.microsoft.com/v1.0/users/${user.id}/joinedTeams`

  const res = await fetch(url, {
    method: 'GET',
    headers,
  });

  const json = await res.json();
  if (json.value) {

    return json.value.map(o => {
      return {
        id: o.id,
        description: o.description,
        teamName: o.displayName,
      };
    });
  }
  return null;
}

const getTeamById = async (data) => {
  const url = `https://graph.microsoft.com/v1.0/teams/${data.id}`
  const headers = {
    Authorization: `Bearer ${data.bearer}`,
    Accept: 'application/json',
  };

  const response = await fetchWithRetry(url, { method: 'GET', headers });
  if (response.ok) {
    return await response.json();
  }

  return null;
};

const getTeamMembers = async (data, env) => {

  const headers = {
    Authorization: `Bearer ${data.bearer}`,
    Accept: 'application/json',
  };

  const url = `https://graph.microsoft.com/v1.0/teams/${data.id}/members`

  const res = await fetch(url, {
    method: 'GET',
    headers,
  });

  const json = await res.json();
  if (json.value) {
    return json.value.map(o => {
      return {
        email: o.email,
        displayName: o.displayName,
        role: o.roles && o.roles.length > 0 ? o.roles[0] : 'unknown',
      };
    });
  }

  return null;
}

const getAllTeams = async (data) => {
  const headers = {
    Authorization: `Bearer ${data.bearer}`,
  };

  const url = `https://graph.microsoft.com/v1.0/teams`;

  const res = await fetch(url, {
    method: 'GET',
    headers,
  });

  await logSearchAttempt({
    searchBy: data.searchBy, // Ensure this is set in data
    searchName: data.nameFilter,
    searchDescription: data.descriptionFilter,
  }, data.env);

  const json = await res.json();
  if (json && json.value) {
    return json.value
    .filter(o => o.visibility !== 'private') // Only public teams
    .filter(o => {
      const name = o.displayName?.toLowerCase() || '';
      const desc = o.description?.toLowerCase() || '';
      return name.includes(data.nameFilter.toLowerCase()) &&
          desc.includes(data.descriptionFilter.toLowerCase());
    })
    .map(o => ({
      id: o.id,
      displayName: o.displayName,
      description: o.description,
    }));
  }

  return null;
};

async function handleMessageStatsRequest(data, env) {
  const teamId = data.body.teamId;

  if (!teamId) {
    return new Response('Missing teamId', { status: 400 });
  }

  try {
    const raw = await env.TEAMS_KV.get(teamId);
    const messageStats = raw ? JSON.parse(raw) : null;
    console.log(`Fetched stats from KV for team ${teamId}:`, messageStats);
    return messageStats;
  } catch (err) {
    console.error(`Error fetching stats for team ${teamId}:`, err);
    return new Response(JSON.stringify({ error: true }), {
      headers: { 'Content-Type': 'application/json' },
      status: 500,
    });
  }
}

async function processInChunks(promises, chunkSize = 5) {
  const results = [];

  for (let i = 0; i < promises.length; i += chunkSize) {
    const chunk = promises.slice(i, i + chunkSize);
    const settled = await Promise.allSettled(chunk);
    results.push(...settled);
  }

  return results;
}

async function getTeamMessageStats(teamId, bearer) {
  const headers = { Authorization: `Bearer ${bearer}` };
  const cutoffDate = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000); // 30 days ago

  const channelsRes = await fetch(`https://graph.microsoft.com/v1.0/teams/${teamId}/channels`, { headers });
  if (!channelsRes.ok) return { messageCount: 0, latestMessage: null, recentCount: 0, partial: true };

  const channels = (await channelsRes.json()).value || [];
  const targetChannel = channels.find(c => ['main', 'general'].includes(c.displayName?.toLowerCase()));
  if (!targetChannel) return { messageCount: 0, latestMessage: null, recentCount: 0 };

  let count = 0;
  let recentCount = 0;
  let latest = null;
  let url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${targetChannel.id}/messages`;

  const allMessages = [];

  while (url) {
    const res = await fetchWithRetry(url, { headers });
    if (!res.ok) break;

    const data = await res.json();
    allMessages.push(...(data.value || []));
    url = data['@odata.nextLink'] || null;
  }

  // Process top-level messages
  const replyFetches = [];

  for (const msg of allMessages) {
    if (msg.from?.user) {
      count++;
      const ts = new Date(msg.lastModifiedDateTime || msg.createdDateTime);
      if (!latest || ts > latest) latest = ts;
      if (ts >= cutoffDate) recentCount++;
    }

    // Instead of nested reply loop, create a fetcher promise
    replyFetches.push(fetchRepliesAndCount(msg.id, headers, teamId, targetChannel.id, cutoffDate));
  }

  const replyResults = await processInChunks(replyFetches, 5);

  for (const result of replyResults) {
    if (result.status === 'rejected') {
      console.error('Reply fetch failed:', result.reason);
    }
  }

  for (const result of replyResults) {
    if (result.status === "fulfilled" && result.value) {
      const { replyCount, recentReplyCount, latestReply } = result.value;
      count += replyCount;
      recentCount += recentReplyCount;
      if (latestReply && (!latest || latestReply > latest)) latest = latestReply;
    }
  }

  return {
    messageCount: count,
    latestMessage: latest ? latest.toISOString().split('T')[0] : null,
    recentCount,
  };
}

async function fetchRepliesAndCount(messageId, headers, teamId, channelId, cutoffDate) {
  let replyCount = 0;
  let recentReplyCount = 0;
  let latestReply = null;
  let url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}/messages/${messageId}/replies`;

  while (url) {
    const res = await fetchWithRetry(url, { headers });
    if (!res.ok) break;

    const data = await res.json();
    const replies = data.value || [];

    for (const reply of replies) {
      if (reply.from?.user) {
        replyCount++;
        const ts = new Date(reply.lastModifiedDateTime || reply.createdDateTime);
        if (!latestReply || ts > latestReply) latestReply = ts;
        if (ts >= cutoffDate) recentReplyCount++;
      }
    }
    url = data['@odata.nextLink'] || null;
    await new Promise(r => setTimeout(r, 500));

  }

  return { replyCount, recentReplyCount, latestReply };
}

async function inviteUser(data, env) {
  const url = `https://graph.microsoft.com/v1.0/invitations`;

  const headers = {
    Authorization: `Bearer ${data.bearer}`,
    'Content-Type': 'application/json',
  };

  const body = {
    "invitedUserEmailAddress": data.body.email,
    "inviteRedirectUrl": "https://teams.microsoft.com",
    "sendInvitationMessage": true,
    "invitedUserDisplayName": data.body.displayName
  }

  const response = await fetch(url, {
    method: 'POST',
    headers,
    body: JSON.stringify(body),
  });

  if (response.ok) {
    return await response.json();
  }
  return null;
}

async function inviteToTeam(data) {
  const url = `https://graph.microsoft.com/v1.0/invitations`;

  const headers = {
    Authorization: `Bearer ${data.bearer}`,
    'Content-Type': 'application/json',
  };

  const body = {
    "invitedUserEmailAddress": data.email,
    "inviteRedirectUrl": `https://teams.microsoft.com/l/team/${data.id}/conversations?groupId=${data.id}&tenantId=36fb8e1d-a891-493f-96d0-3038e4e9291c`,
    "sendInvitationMessage": true,
    "invitedUserDisplayName": data.displayName,
    "invitedUserMessageInfo": {
      "customizedMessageBody": `Hi ${data.displayName},\n\nYou've been invited to join ${data.teamName} Microsoft Team. Click below to accept and join:\n\nhttps://teams.microsoft.com/l/team/${data.id}/conversations?groupId=${data.id}&tenantId=36fb8e1d-a891-493f-96d0-3038e4e9291c`,
    }
  }

  const response = await fetch(url, {
    method: 'POST',
    headers,
    body: JSON.stringify(body),
  });

  if (response.ok) {
    return await response.json();
  }
  return null;
}

// Invite guest if not in directory, else retrieve existing user
async function ensureGuestUser(data) {
  const user = await getUser(data.email, data.bearer);
  if (!user)  {
    console.log("User not found, inviting to team", data.email);
    const invite = await inviteToTeam(data);
    if (invite) {
      return invite.invitedUser.id;
    }
  }

  if (user?.id) return user.id;
  return null;
}

async function addGuestToTeam(data) {
  const headers = {
    Authorization: `Bearer ${data.bearer}`,
    'Content-Type': 'application/json',
  };
  const url = `https://graph.microsoft.com/v1.0/groups/${data.id}/members/$ref`;
  const body = {
    '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${data.userId}`
  }
  return await fetch(url, {
    method: 'POST',
    headers,
    body: JSON.stringify(body),
  });
}

async function addTeamMembers(data, env) {
  const results = [];

  const team = await getTeamById(data);
  if (!team) {
    console.log("Team not found", data.id);
    return results;
  }
  data.teamName = team.displayName || '';
  // Loop over the full user objects: { displayName, email }
  const uniqueUsers = Array.from(new Map(data.body.users.map(u => [u.email, u])).values());
  for (const user of uniqueUsers) {
    const { displayName , email } = user;

    const userId = await ensureGuestUser({...data, email, displayName });
    let added = false;

    if (userId) {
      const response = await addGuestToTeam({ ...data, userId });
      if (response.status === 400 || response.status === 204) {
        added = true;
      }
    }

    results.push({ email, added });

    // Log the member addition
    await logMemberAddition({
      addedBy: data.body.addedBy, // Ensure this is set in data
      addedUser: email,
      teamName: data.teamName,
      added,
    }, env);
  }

  return results;
}

async function removeTeamMembers(data, env) {
  const { id: teamId, body, bearer } = data;
  const results = [];

  const team = await getTeamById(data);
  if (!team) {
    console.log("Team not found", teamId);
    return results;
  }

  for (const user of body.users) {
    const userObj = await getUser(user.email, bearer);
    if (!userObj) {
      results.push({ email: user.email, removed: false, reason: 'User not found' });
      continue;
    }
    const url = `https://graph.microsoft.com/v1.0/groups/${teamId}/members/${userObj.id}/$ref`;
    const headers = {
      Authorization: `Bearer ${bearer}`,
      Accept: 'application/json',
    };

    const res = await fetch(url, { method: 'DELETE', headers });
    results.push({ email: user.email, removed: res.ok });

    // Log the member removal
    await logMemberRemoval({
      removedBy: data.body.removedBy,
      removedUser: user.email,
      teamName: team.displayName,
      removed: res.ok,
    }, env);
  }

  return results;
}

async function fetchWithRetry(url, options = {}, retries = 4, delay = 5000, timeout = 10000) {
  for (let i = 0; i < retries; i++) {
    const controller = new AbortController();
    const id = setTimeout(() => controller.abort(), timeout);
    options.signal = controller.signal;

    try {
      const res = await fetch(url, options);
      clearTimeout(id);

      if (res.ok) return res;

      const body = await res.text();
      const retryAfter = res.headers.get("Retry-After");

      console.warn(`Retry ${i + 1}/${retries}: ${res.status} - ${body}`);

      if ([502, 503, 504, 429].includes(res.status)) {
        const baseDelay = retryAfter ? parseInt(retryAfter) * 1000 : delay;
        const jitter = Math.floor(Math.random() * 1000); // up to 1s of jitter
        await new Promise(r => setTimeout(r, baseDelay + jitter));
        delay *= 2;
        continue;
      }

      throw new Error(`Non-retryable HTTP error: ${res.status} - ${body}`);
    } catch (err) {
      clearTimeout(id);

      if (err.name === "AbortError") {
        console.warn(`Timeout on attempt ${i + 1} (${timeout}ms)`);
      } else {
        console.warn(`Fetch error on attempt ${i + 1}:`, err.message);
      }

      if (i < retries - 1) {
        const jitter = Math.floor(Math.random() * 1000);
        await new Promise(r => setTimeout(r, delay + jitter));
        delay *= 2;
        continue;
      }

      throw new Error(`Failed after ${retries} retries: ${err.message}`);
    }
  }
}



export {
  getUserTeams,
  addTeamMembers,
  removeTeamMembers,
  getTeamMembers,
  getTeamById,
  getAllTeams,
  handleMessageStatsRequest,
  getTeamMessageStats,
  inviteUser
}
