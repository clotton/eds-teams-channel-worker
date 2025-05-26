import pLimit from 'p-limit';
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

  if (res.status === 404) {
    return { notFound: true };
  }

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

async function getTeamMessageStats(data) {
  const headers = { Authorization: `Bearer ${data.bearer}` };
  const teamId = data.body.teamIds;
  const now = new Date();
  const cutoffDate = new Date(now.getTime() - 30 * 24 * 60 * 60 * 1000); // 30 days ago

  console.log(`Fetching channels for team ${teamId}`);
  const channelsRes = await fetchWithRetry(`https://graph.microsoft.com/v1.0/teams/${teamId}/channels`, { headers });

  if (!channelsRes.ok) {
    console.error(`Failed to fetch channels: ${channelsRes.status} ${channelsRes.statusText}`);
    return { count: 0, latestMessageDate: null, recentCount: 0 };
  }

  const channels = (await channelsRes.json()).value || [];

  const results = [];

  for (const channel of channels) {
    let count = 0;
    let recentCount = 0;
    let latest = null;
    let url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channel.id}/messages`;

    while (url) {
      console.log(`Fetching messages from: ${url}`);
      const res = await fetchWithRetry(url, { headers });

      if (!res.ok) {
        const errorText = await res.text();
        console.error(`Error fetching messages: ${res.status} ${res.statusText}`);
        console.error(`URL: ${url}`);
        console.error(`Response: ${errorText}`);
        break;
      }

      const data = await res.json();
      const messages = data.value || [];

      for (const msg of messages) {
        if (msg.from?.user) {
          count++;
          const ts = new Date(msg.lastModifiedDateTime || msg.createdDateTime);
          if (!latest || ts > latest) latest = ts;
          if (ts >= cutoffDate) recentCount++;
        }

        let replyUrl = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channel.id}/messages/${msg.id}/replies`;
        while (replyUrl) {
          console.log(`Fetching replies from: ${replyUrl}`);
          const replyRes = await fetchWithRetry(replyUrl, { headers });

          if (!replyRes.ok) {
            const replyErrorText = await replyRes.text();
            console.error(`Error fetching replies: ${replyRes.status} ${replyRes.statusText}`);
            console.error(`URL: ${replyUrl}`);
            console.error(`Response: ${replyErrorText}`);
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

          replyUrl = replyData['@odata.nextLink'];
        }
      }

      url = data['@odata.nextLink'];
    }

    results.push({ count, latest, recentCount });
  }

  const totalCount = results.reduce((sum, r) => sum + r.count, 0);
  const recentCount = results.reduce((sum, r) => sum + r.recentCount, 0);
  const latestMessageDate = results.reduce((latest, r) =>
          !latest || (r.latest && r.latest > latest) ? r.latest : latest,
      null
  );

  return { messageCount: totalCount, latestMessage: latestMessageDate, recentCount };
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
  if (user?.notFound) {
    console.log("User not found, inviting to team", data.email);
   const invite = await inviteToTeam(data);
    if (invite) {
      return invite.invitedUser.id;
    }
  }
  console.log("User found", user);
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
  for (const user of data.body) {
    const { email, displayName } = user;

    const userId = await ensureGuestUser({...data, email, displayName });
    let added = false;

    if (userId) {
      const response = await addGuestToTeam({ ...data, userId });

      if (response.status === 400) {
        console.log("User already in team", email);
      }

      if (response.status === 204) {
        added = true;
      }
    }

    results.push({ email, added });
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
  getTeamMessageStats,
  inviteUser
}
