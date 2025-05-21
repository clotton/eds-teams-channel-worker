
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

  const json = await response.json();
  if (json.value) {
    return json;
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

const asyncForEach = async (array, callback) => {
  for (let index = 0; index < array.length; index++) {
    await callback(array[index]);
  }
}

const addRemoveUserToTeams = async (data) => {
  const user = await getUser(data.id, data.bearer);
  if (!user) return null;

  const result = {
    add: {
      success: [],
      failed: [],
    },
    remove: {
      success: [],
      failed: [],
    }
  };

  await asyncForEach(data.body.add, async (id) => {
    const team = await getTeamById({id, bearer: data.bearer});
    if (team) {
      const headers = {
        Authorization: `Bearer ${data.bearer}`,
        'Content-Type': 'application/json',
      };

      const url = `https://graph.microsoft.com/v1.0/groups/${team.id}/members/$ref`;
      const body = {
        '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${user.id}`
      }
      const res = await fetch(url, {
          method: 'POST',
          headers,
          body: JSON.stringify(body),
        });
        if (res.status === 204) {
          result.add.success.push(id);
        } else {
          result.add.failed.push(id);
        }

    }
  });

  await asyncForEach(data.body.remove, async (id) => {
    const team = await getTeamById({id, bearer: data.bearer});
    if (team) {
      const headers = {
        Authorization: `Bearer ${data.bearer}`,
        'Content-Type': 'application/json',
      };

      const url = `https://graph.microsoft.com/v1.0/groups/${team.id}/members/${user.id}/$ref`;
        const res = await fetch(url, {
          method: 'DELETE',
          headers,
        });
        if (res.status === 204) {
          result.remove.success.push(id);
        } else {
          result.remove.failed.push(id);
        }
    }
  });

  return result;
};

async function getTotalTeamMessages(data) {
  const headers = { Authorization: `Bearer ${data.bearer}` };
  const channelsRes = await fetch(`https://graph.microsoft.com/v1.0/teams/${data.id}/channels`, { headers });

  if (!channelsRes.ok) return { count: 0, latestMessageDate: null };

  const channels = (await channelsRes.json()).value || [];

  const results = await Promise.all(channels.map(async (channel) => {
    let count = 0;
    let latest = null;
    let url = `https://graph.microsoft.com/v1.0/teams/${data.id}/channels/${channel.id}/messages`;

    while (url) {
      const res = await fetch(url, { headers });
      if (!res.ok) break;

      const data = await res.json();
      const messages = data.value || [];
      count += messages.filter(msg => msg.from?.user).length; // Only count human messages

      // Track the latest message date
      for (const msg of messages) {
        const ts = msg.lastModifiedDateTime || msg.createdDateTime;
        if (msg.from?.user && (!latest || ts > latest)) {
          latest = ts;
        }
      }

      url = data['@odata.nextLink'];
    }

    return { count, latest };
  }));

  let totalCount = 0;
  let globalLatest = null;

  for (const r of results) {
    totalCount += r.count;
    if (!globalLatest || (r.latest && r.latest > globalLatest)) {
      globalLatest = r.latest;
    }
  }

  return globalLatest ? { count: totalCount, latestMessageDate: globalLatest.split('T')[0] } : { count: 0, latestMessageDate: null };
}

async function inviteGuest(data) {
  const url = `https://graph.microsoft.com/v1.0/invitations`;

  const headers = {
    Authorization: `Bearer ${data.bearer}`,
    'Content-Type': 'application/json',
  };

  const body = {
    "invitedUserEmailAddress": data.body.email,
    "inviteRedirectUrl": "https://teams.microsoft.com",
    "sendInvitationMessage": true,
    "invitedUserDisplayName": data.body.name
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
  const user = await xgetUser(data.id, data.bearer);
  if (user?.notFound){
    const invite = await inviteGuest(data );

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
  for (const email of data.body.guests) {
    const userId = await ensureGuestUser({ id: email, bearer: data.bearer });
    let added = false;
    if (userId) {
      const response = await addGuestToTeam({ ...data, userId });
      if (response.status === 204) {
        added = true;
      }
    }
    results.push({ email, added });
  }
  return results;
}

export {
  getUserTeams,
  addTeamMembers,
  getTeamMembers,
  getTeamById,
  getAllTeams,
  getTotalTeamMessages,
  addRemoveUserToTeams,
  inviteGuest
}
