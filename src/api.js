import logo from './logo.js';
import { isQuestion, stripHtml } from "./utils";

export async function logEvent(message, env) {
  const webhookUrl = env.SLACK_WEBHOOK_URL;
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

const getOwners = async (bearer) => {
  const params = new URLSearchParams({
    '$filter': `startsWith(mail,'owner_')`,
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

  if (res.ok) {
    const json = await res.json();
    if (json.value && json.value.length > 0) {
      return json.value.map(o => {
        return {
          id: o.id,
          email: o.mail,
          displayName: o.displayName,
        };
      });
    }
  }

  return [];
};

const updateTeamPhoto = async (data) => {
  const { id } = data.body;
  if (id) {
    const headers = {
      Authorization: `Bearer ${data.bearer}`,
      'Content-Type': 'image/png',
    };
    const url = `https://graph.microsoft.com/v1.0/groups/${id}/photo/$value`;
    console.log('Updating photo', url);

    const res = await fetch(url, {
      method: 'PUT',
      headers,
      body: logo(),
    });

    console.log('Photo updated', res.status, res.statusText);
  }
};

const addOwnersToTeam = async (teamId, owners, bearer) => {

    const results = [];

	for (const owner of owners) {
	    if (!owner?.id) {
          console.warn(`⚠️ Skipping owner without id:`, owner);
          results.push({
            id: owner?.id,
            email: owner?.email,
            ok: false,
            error: 'missing id'
          });
          continue;
        }

		const url = `https://graph.microsoft.com/v1.0/groups/${teamId}/owners/$ref`;
		const body = {
			"@odata.id": `https://graph.microsoft.com/v1.0/directoryObjects/${owner.id}`
		};

		try {
			const response = await fetch(url, {
				method: "POST",
				headers: {
					Authorization: `Bearer ${bearer}`,
					"Content-Type": "application/json"
				},
				body: JSON.stringify(body)
			});

			if (response.ok) {
				console.log(`✅ Added owner ${owner.email || owner.id} to team ${teamId}`);
          results.push({
            id: owner.id,
            email: owner.email,
            ok: true
          });
			} else {
				  const error = await response.json().catch(() => ({}));
				console.log(
                          `❌ Failed to add owner ${owner.email || owner.id} to team ${teamId}:`,
                          error
                        );
          results.push({
            id: owner.id,
            email: owner.email,
            ok: false,
            error
          });
			}
		} catch (err) {
		console.error(
                `⚠️ Error adding owner ${owner.email || owner.id} to team ${teamId}:`,
                err
              );
        results.push({
          id: owner.id,
          email: owner.email,
          ok: false,
          error: err.message || String(err)
        });
		}
	}

  return results;
};

const getChannels = async (teamId, bearer) => {
  const headers = {
    Authorization: `Bearer ${bearer}`,
  };

  const url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels`;

  const res = await fetch(url, {
    method: 'GET',
    headers,
  });

  const json = await res.json();
  if (json && json.value) {
    return json.value;
  }

  return null;
}

const renameChannel = async (teamId, channelId, bearer) => {
  const headers = {
    Authorization: `Bearer ${bearer}`,
    'Content-Type': 'application/json',
  };

  const url = `https://graph.microsoft.com/v1.0/teams/${teamId}/channels/${channelId}`;
  const body = {
      "displayName": 'Main',
    }

  const res = await fetch(url, {
    method: 'PATCH',
    headers,
    body: JSON.stringify(body),
  });

  console.log('Channel renamed', res.status, res.statusText);

};

const createTeam = async (data, env) => {
  const { createdBy, name, teamType, description = '' } = data.body;

  if (name) {
    console.log('Creating team', name, description);
    const headers = {
      Authorization: `Bearer ${data.bearer}`,
      'Content-Type': 'application/json',
    };
    const url = `https://graph.microsoft.com/v1.0/teams`;
    const body = {
      'template@odata.bind': 'https://graph.microsoft.com/v1.0/teamsTemplates(\'standard\')',
      visibility: 'public',
      displayName: name,
      description,
      guestSettings: {
        'allowCreateUpdateChannels': true,
      },
      members:[]
    };
    const owners = await getOwners(data.bearer);
    if (!owners || owners.length === 0) {
      console.error('No owners found');
      return null;
    }
    // api accepts only 1 member...
    owners.filter(o => o.email.startsWith('owner_ck')).forEach(o => {
      body.members.push({
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles:[
          'owner'
        ],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${o.id}')`
      });
    });

    // 1. create team with initial owner
    const res = await fetch(url, { method: 'POST', headers, body: JSON.stringify(body) });
    if (!res.ok) throw new Error('Failed to create team');
    const location = res.headers.get('location');
    const id = location.split("'")[1];

    // 2. wait for team to be ready (polling)
    let ready = false, retries = 10;
    while (!ready && retries-- > 0) {
      await new Promise(r => setTimeout(r, 2000));
      const team = await getTeamById({ id, bearer: data.bearer });
      if (team) ready = true;
    }
    if (!ready) throw new Error('Team provisioning timeout');
    console.log('Team created', id);

    // 3. update photo
    await updateTeamPhoto({ bearer: data.bearer, body: { id } });

    // 4. rename channel to 'Main' (if exists)
    const channels = await getChannels(id, data.bearer);
    const targetChannel = channels?.find(c => c.displayName?.toLowerCase() === 'general');
    if (targetChannel) await renameChannel(id, targetChannel.id, data.bearer);

    // 5. add remaining owners
    const remaining = owners
    .filter(o => !o.email.startsWith('owner_ck'))
    .map(o => ({ id: o.id, email: o.email }));
    const ownerAddResults = await addOwnersToTeam(id, remaining, data.bearer);

    const failedOwners = ownerAddResults.filter(r => !r.ok);
    const failedOwnerSummary = failedOwners.map(o =>
      `${o.email || o.id} - ${typeof o.error === 'string'
        ? o.error
        : JSON.stringify(o.error).slice(0, 200)}`
    ).join('\n');

    // 6. add guests
    const teamMembers = (teamType === 'EDS' ? env.EDS_GUESTS : env.LLMO_GUESTS).split(',').map(e => e.trim()).filter(Boolean);
    const users = await Promise.all(teamMembers.map(email => getUser(email, data.bearer)));
    const validUsers = users.filter(Boolean).map(u => ({ id: u.id, email: u.mail }));

    const guestResults = [];
    let count = 0;
    for (const u of validUsers) {
      const res = await addGuestToTeam({ id, bearer: data.bearer, userId: u.id });
      const ok = res.ok || res.status === 204 || res.status === 400;
      if (ok) count = count + 1;
      guestResults.push({
        email: u.email,
        ok,
        status: res.status,
      });
    }
    console.log(`Added guests:`, count);

    const failedGuests = guestResults.filter(g => !g.ok);
    const failedGuestSummary = failedGuests
      .map(g => `${g.email || 'unknown'} - HTTP ${g.status}`)
      .join('\n');

    // 7.  create the admin tag
    await createAdminTag(id, owners.map(o => o.id), data.bearer);

    // 8.  log team creation event
    await logEvent({
      text:
        `👤 *${createdBy}* created team *${name}* (type: *${teamType || 'N/A'}*) — ${count} guests added` +
        (failedOwners.length
          ? `\n⚠️ Failed to add ${failedOwners.length} owner(s):\n${failedOwnerSummary}`
          : '') +
        (failedGuests.length
          ? `\n⚠️ Failed to add ${failedGuests.length} guest(s):\n${failedGuestSummary}`
          : '')
    }, env);

    return {
      name,
      description,
    };
  }
  return null;
};

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
        id: o.id,
        email: o.email,
        displayName: o.displayName,
        role: o.roles && o.roles.length > 0 ? o.roles[0] : 'unknown',
      };
    });
  }

  return null;
}

const getAllTeams = async (data) => {
  const { searchBy, nameFilter = '*', descriptionFilter = '*', env, bearer } = data;

  const headers = {
    Authorization: `Bearer ${bearer}`,
  };

  let url = `https://graph.microsoft.com/v1.0/teams?$top=100`;
  const results = [];

  const nameFilterLower = nameFilter.toLowerCase();
  const descFilterLower = descriptionFilter.toLowerCase();

  // Log the search event
  await logEvent({
    text: `👤 *${searchBy}* searched Teams for name: *${nameFilter}* and description: *${descriptionFilter}*`,
  }, env);

  while (url) {
    const res = await fetch(url, {
      method: 'GET',
      headers,
    });

    if (!res.ok) {
      console.error(`❌ Failed to fetch teams: ${res.status} ${res.statusText}`);
      break;
    }

    const json = await res.json();
    const teams = json.value || [];

    // Filter as you go to reduce memory usage
    for (const o of teams) {
      const name = o.displayName?.toLowerCase() || '';
      const desc = o.description?.toLowerCase() || '';

      const matches =
        o.visibility !== 'private' &&
        name.includes(nameFilterLower) &&
        desc.includes(descFilterLower);

      if (matches) {
        results.push({
          id: o.id,
          displayName: o.displayName,
          description: o.description,
        });
      }
    }

    url = json['@odata.nextLink'] || null;
  }

  return results;
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

async function processInChunks(teamId, tasks, concurrency = 1, maxSubrequests = 1000) {
  const results = [];
  const executing = [];
  let subrequestCount = 0;

  for (const task of tasks) {
    const wrappedTask = async () => {
      if (subrequestCount >= maxSubrequests) {
        throw new Error(`❌ Max subrequest count (${maxSubrequests}) reached for ${teamId}. Skipping remaining tasks.`);
      }
      try {
        const result = await task();
        subrequestCount++; // assume 1 subrequest per task; adjust if task uses more
        return { status: "fulfilled", value: result };
      } catch (err) {
        return { status: "rejected", reason: err };
      }
    };

    const p = wrappedTask();
    results.push(p);

    if (concurrency <= tasks.length) {
      const e = p.then(() => executing.splice(executing.indexOf(e), 1));
      executing.push(e);
      if (executing.length >= concurrency) {
        await Promise.race(executing);
      }
    }
  }

  return Promise.all(results);
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
  let questionCount = 0;
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

      const plainText = stripHtml(msg.body?.content);

      if (isQuestion(plainText)) {
        questionCount++;
      }
    }

    replyFetches.push(() => fetchRepliesAndCount(msg.id, headers, teamId, targetChannel.id, cutoffDate));
  }
  const replyResults = await processInChunks(teamId, replyFetches);

  for (const result of replyResults) {
    if (result.status === 'rejected') {
      console.error('Reply fetch failed:', result.reason);
    }
  }

  for (const result of replyResults) {
    if (result.status === "fulfilled" && result.value) {
      const {
        replyCount,
        recentReplyCount,
        latestReply,
        replyQuestionCount,
      } = result.value;

      count += replyCount;
      recentCount += recentReplyCount;
      questionCount += replyQuestionCount;

      if (latestReply && (!latest || latestReply > latest)) latest = latestReply;
    }
  }


  return {
    messageCount: count,
    latestMessage: latest ? latest.toISOString().split('T')[0] : null,
    recentCount,
    questionCount,
  };
}

async function fetchRepliesAndCount(messageId, headers, teamId, channelId, cutoffDate) {
  let replyCount = 0;
  let recentReplyCount = 0;
  let replyQuestionCount = 0;
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

        const plainText = stripHtml(reply.body?.content);

        if (isQuestion(plainText)) {
            replyQuestionCount++;
        }
      }
    }
    url = data['@odata.nextLink'] || null;
    await new Promise(r => setTimeout(r, 500));

  }

  return {
    replyCount,
    recentReplyCount,
    latestReply,
    replyQuestionCount,
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
async function ensureGuestUser(data){
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

  const uniqueUsers = Array.from(new Map(data.body.users.map(u => [u.email, u])).values());
  for (const user of uniqueUsers) {
    const { displayName, email } = user;
    const userId = await ensureGuestUser({...data, email, displayName });
    let added = false;

    if (userId) {
      const response = await addGuestToTeam({ ...data, userId });
      if (response.status === 400 || response.status === 204) {
        added = true;
      }
    }

    results.push({ email, added });

    await logEvent({
      text: `👤 *${data.body.addedBy}* added *${email}* to team *${team.displayName}* — ${added ? '✅ *Success*' : '❌ *Failed*'}`
      }, env);
  }

  return results;
}

async function removeTeamMembers(data, env) {
  const results = [];

  const team = await getTeamById(data);
  if (!team) {
    console.log("Team not found", data.id);
    return results;
  }

  for (const user of data.body.users) {
    const userObj = await getUser(user.email, data.bearer);
    if (!userObj) {
      results.push({ email: user.email, removed: false, reason: 'User not found' });
      continue;
    }

    const url = `https://graph.microsoft.com/v1.0/groups/${data.id}/members/${userObj.id}/$ref`;
    const headers = {
      Authorization: `Bearer ${data.bearer}`,
      Accept: 'application/json',
    };

    const res = await fetch(url, { method: 'DELETE', headers });
    results.push({ email: user.email, removed: res.ok });

    await logEvent({
        text: `👤 *${data.body.removedBy}* removed *${user.email}* to team *${team.displayName}* — ${res.ok
          ? '✅ Success' : '❌ Failed'}`
      }, env);
  }

  return results;
}

async function fetchWithRetry(url, options = {}, retries = 4, delay = 5000, timeout = 20000) {
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
      console.warn(`Retry ${i + 1}/${retries}: ${res.status} - ${body} [${url}]`);

      if (options.retry !== false && [502, 503, 504, 429].includes(res.status)) {
        const baseDelay = retryAfter ? parseInt(retryAfter) * 1000 : delay;
        const jitter = Math.floor(Math.random() * 1000);
        await new Promise(r => setTimeout(r, baseDelay + jitter));
        delay = Math.min(delay * 2, 20000);
        continue;
      }

      throw new Error(`Non-retryable HTTP error: ${res.status} - ${body}`);
    } catch (err) {
      clearTimeout(id);

      if (err.name === "AbortError") {
        console.warn(`Timeout on attempt ${i + 1} (${timeout}ms) for ${url}`);
      } else {
        console.warn(`Fetch error on attempt ${i + 1} for ${url}: ${err.message}`);
      }

      if (i < retries - 1) {
        const jitter = Math.floor(Math.random() * 1000);
        await new Promise(r => setTimeout(r, delay + jitter));
        delay = Math.min(delay * 2, 20000);
        continue;
      }

      throw new Error(`Failed after ${retries} retries: ${err.message}`);
    }
  }
}

async function createAdminTag(teamId, userIds, token) {
  const url = `https://graph.microsoft.com/v1.0/teams/${teamId}/tags`;

  const body = {
    displayName: "admin",
    members: userIds.map(userId => ({ userId }))
  };

  try {
    const res = await fetch(url, {
      method: "POST",
      headers: {
        "Authorization": `Bearer ${token}`,
        "Content-Type": "application/json"
      },
      body: JSON.stringify(body)
    });

    if (!res.ok) {
      const errorText = await res.text();
      console.error(`Failed to create admin tag: ${res.status} - ${errorText}`);
      return null;
    }

    const data = await res.json();
    console.log("Admin tag created:", data);
    return data;
  } catch (err) {
    console.error("Error creating admin tag:", err);
    return null;
  }
}

export {
  getUserTeams,
  createTeam,
  addTeamMembers,
  removeTeamMembers,
  getTeamMembers,
  getTeamById,
  getAllTeams,
  handleMessageStatsRequest,
  getTeamMessageStats,
  inviteUser
}
