export async function logMemberAddition({ addedBy, addedUser, teamName, added }, env) {
  const webhookUrl = env.SLACK_WEBHOOK_URL; // Replace with your webhook
  const message = {
    text: `👤 *${addedBy}* attempted to add *${addedUser}* to team *${teamName}* — ${added
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

const getUserTeams = async (data) => {
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

  const response = await fetch(url, { method: 'GET', headers });
  if (response.ok) {
    return await response.json();
  }

  return null;
};

const getTeamMembers = async (data) => {

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
    return await env.TEAMS_KV.get(teamId);
  } catch (err) {
    console.error(`Error fetching stats for team ${teamId}:`, err);
    return new Response(JSON.stringify({ error: true }), {
      headers: { 'Content-Type': 'application/json' },
      status: 500,
    });
  }
}


function pLimit(concurrency) {
  const queue = [];
  let activeCount = 0;

  const next = () => {
    if (queue.length === 0 || activeCount >= concurrency) return;
    activeCount++;
    const { fn, resolve, reject } = queue.shift();
    fn().then(resolve).catch(reject).finally(() => {
      activeCount--;
      next();
    });
  };

  return (fn) => new Promise((resolve, reject) => {
    queue.push({ fn, resolve, reject });
    next();
  });
}

async function getTeamMessageStats(teamId, bearer) {
  const headers = { Authorization: `Bearer ${bearer}` };
  const cutoffDate = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);

  const channelsRes = await fetch(`https://graph.microsoft.com/v1.0/teams/${teamId}/channels`, { headers });
  if (!channelsRes.ok) return { messageCount: 0, latestMessage: null, recentCount: 0 };

  const channels = (await channelsRes.json()).value || [];
  const targetChannel = channels.find(c => c.displayName.toLowerCase() === 'main')
      || channels.find(c => c.displayName.toLowerCase() === 'general');
  if (!targetChannel) return { messageCount: 0, latestMessage: null, recentCount: 0 };

  let count = 0;
  let recentCount = 0;
  let latest = null;

  // Step 1: Fetch all message pages
  const allMessages = [];
  let url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${targetChannel.id}/messages`;
  while (url) {
    const res = await fetchWithRetry(url, { headers });
    if (!res.ok) {
      console.error(`Error fetching messages page: ${res.status}`);
      break;
    }
    const data = await res.json();
    allMessages.push(...(data.value || []));
    url = data['@odata.nextLink'] || null;
  }

  // Step 2: Prepare throttled fetch for replies
  const limit = pLimit(10); // 10 concurrent fetches
  const replyTasks = [];

  for (const msg of allMessages) {
    if (msg.from?.user) {
      count++;
      const ts = new Date(msg.lastModifiedDateTime || msg.createdDateTime);
      if (!latest || ts > latest) latest = ts;
      if (ts >= cutoffDate) recentCount++;
    }

    // Fetch replies
    const repliesUrl = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${targetChannel.id}/messages/${msg.id}/replies`;

    replyTasks.push(limit(async () => {
      let replyUrl = repliesUrl;
      while (replyUrl) {
        try {
          const replyRes = await fetchWithRetry(replyUrl, { headers });
          if (!replyRes.ok) {
            const errorText = await replyRes.text();
            console.error(`Failed to fetch replies for ${msg.id}: ${replyRes.status} - ${errorText}`);
            break;
          }

          const replyData = await replyRes.json();
          const replies = replyData.value || [];

          for (const reply of replies) {
            if (reply.from?.user) {
              count++;
              const ts = new Date(reply.lastModifiedDateTime || reply.createdDateTime);
              if (!latest || ts > latest) latest = ts;
              if (ts >= cutoffDate) recentCount++;
            }
          }

          replyUrl = replyData['@odata.nextLink'] || null;
        } catch (err) {
          console.error(`Error fetching replies for ${msg.id}:`, err);
          break;
        }
      }
    }));
  }

  // Step 3: Run all throttled reply fetches in parallel
  await Promise.allSettled(replyTasks);

  return {
    messageCount: count,
    latestMessage: latest ? latest.toISOString().split('T')[0] : null,
    recentCount
  };
}

async function inviteUser(data) {
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

async function addTeamMembers(data) {
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
    }, data.env);
  }

  return results;
}

async function fetchWithRetry(url, options, retries = 5) {
  for (let i = 0; i < retries; i++) {
    const res = await fetch(url, options);

    if (res.status !== 429) return res;

    const retryAfter = res.headers.get("Retry-After");
    const waitTime = retryAfter ? parseInt(retryAfter, 10) * 1000 : 5000; // default 5s
    console.warn(`Rate limited. Waiting ${waitTime / 1000}s before retrying... [${i + 1}/${retries}]`);

    await new Promise(resolve => setTimeout(resolve, waitTime));
  }

  throw new Error(`Failed after ${retries} retries due to 429 errors: ${url}`);
}


export {
  getUserTeams,
  addTeamMembers,
  getTeamMembers,
  getTeamById,
  getAllTeams,
  handleMessageStatsRequest,
  getTeamMessageStats,
  inviteUser
}
