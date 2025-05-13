
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

  try {
    const response = await fetch(url, { method: 'GET', headers });
    if (!response.ok) {
      console.warn(`Error fetching team ${id}: ${response.statusText}`);
      return null;
    }
    return response.json();
  } catch (error) {
    console.error(`Failed to fetch team ${data.id}:`, error);
    return null;
  }
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
    const team = await getTeamById(id, data.bearer);

    if (team) {
      const headers = {
        Authorization: `Bearer ${data.bearer}`,
        'Content-Type': 'application/json',
      };

      const url = `https://graph.microsoft.com/v1.0/groups/${team.id}/members/$ref`;
      const body = {
        '@odata.id': `https://graph.microsoft.com/v1.0/directoryObjects/${user.id}`
      }
      try {
        const res = await fetch(url, {
          method: 'POST',
          headers,
          body: JSON.stringify(body),
        });

        if (res.status === 204) {
          result.add.success.push(id);
        } else {
          console.error('Failed to add user to team:', id, res.status,
              res.statusText);
          result.add.failed.push(id);
        }
      } catch (err) {
        console.error('Error adding user to team:', id, err);
        result.add.failed.push(id);
      }
    }
  });

  await asyncForEach(data.body.remove, async (id) => {
    const team = await getTeamById(id, data.bearer);

    if (team) {
      const headers = {
        Authorization: `Bearer ${data.bearer}`,
        'Content-Type': 'application/json',
      };

      const url = `https://graph.microsoft.com/v1.0/groups/${team.id}/members/${user.id}/$ref`;
      try {
        const res = await fetch(url, {
          method: 'DELETE',
          headers,
        });

        if (res.status === 204) {
          result.remove.success.push(id);
        } else {
          result.remove.failed.push(id);
        }
      } catch (err) {
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

  const trimmedDate = globalLatest.split('T')[0];

  return {
    messageCount: totalCount,
    latestMessageDate: trimmedDate
  };
}

export {
  getUserTeams,
  getTeamMembers,
  getTeamById,
  getAllTeams,
  getTotalTeamMessages,
  addRemoveUserToTeams,
}
