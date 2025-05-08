
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

const getTeam = async (displayName, bearer) => {
  const params = new URLSearchParams({
    '$filter': `(displayName eq '${displayName}')`,
    '$select': 'id,displayName,createdDateTime',
  });

  const headers = {
    ConsistencyLevel: 'eventual',
    Authorization: `Bearer ${bearer}`,
  };

  const url = `https://graph.microsoft.com/v1.0/groups?${params}`

  const res = await fetch(url, {
    method: 'GET',
    headers,
  });

  const json = await res.json();
  if (json.value && json.value.length > 0) {
    return json.value[0];
  }

  return null;
}

const getTeamById = async (id, bearer) => {
  const headers = {
    ConsistencyLevel: 'eventual',
    Authorization: `Bearer ${bearer}`,
  };

  const url = `https://graph.microsoft.com/v1.0/teams/${id}`

  const res = await fetch(url, {
    method: 'GET',
    headers,
  });

  return res.json();
};

const getTeamMembers = async (data) => {
  let { id, name, bearer } = data;
  if (!id && name) {
    const team = await getTeam(name, data.bearer);
    if (!team) return null;
    id = team.id;
  }

  const headers = {
    ConsistencyLevel: 'eventual',
    Authorization: `Bearer ${bearer}`,
  };

	  const url = `https://graph.microsoft.com/v1.0/teams/${id}/members`

  const res = await fetch(url, {
    method: 'GET',
    headers,
  });

  const json = await res.json();
  if (json.value) {
    return json.value.map(o => {
      return {
        email: o.email,
        teamName: o.displayName,
        role: o.roles && o.roles.length > 0 ? o.roles[0] : 'unknown',
      };
    });
  }

  return null;
}

const getAllTeams = async (accessToken, nameFilter = '', descriptionFilter = '') => {
  const url = `https://graph.microsoft.com/v1.0/teams`;
  const response = await fetch(url, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json',
    },
  });

  if (!response.ok) return response;

  const json = await response.json();

  return json.value || [];
};

const asyncForEach = async (array, callback) => {
  for (let index = 0; index < array.length; index++) {
    await callback(array[index]);
  }
}

const addRemoveUserToTeams = async (email, body, bearer) => {
  const user = await getUser(email, bearer);
  console.log('user:', user);
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

  await asyncForEach(body.add, async (id) => {
    console.log('Adding to team ID:', id);
    const team = await getTeamById(id, bearer);
    console.log('team:', team);

    if (team) {
      const headers = {
        Authorization: `Bearer ${bearer}`,
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

  await asyncForEach(body.remove, async (id) => {
    console.log('Removing from team ID:', id);
    const team = await getTeamById(id, bearer);
    console.log('team:', team);

    if (team) {
      const headers = {
        Authorization: `Bearer ${bearer}`,
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
          console.error('Failed to remove user from team:', id, res.status,
              res.statusText);
          result.remove.failed.push(id);
        }
      } catch (err) {
        console.error('Error removing user from team:', id, err);
        result.remove.failed.push(id);
      }
    }
  });

  return result;
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

const getOwners = async (bearer) => {
  const params = new URLSearchParams({
    '$filter': `startsWith(mail,'admin_')`,
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

  if (res.status === 200) {
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

const addMembers = async (teamId, members, bearer) => {
  const headers = {
    Authorization: `Bearer ${bearer}`,
    'Content-Type': 'application/json',
  };

  const url = `https://graph.microsoft.com/v1.0/teams/${teamId}/members/add`;
  const body = {
    values:[]
  };

  members.forEach(member => {
    const m = {
      '@odata.type': 'microsoft.graph.aadUserConversationMember',
      roles: [],
      'user@odata.bind': `https://graph.microsoft.com/v1.0/users(\'${member.id}\')`
    };
    if (member.role) m.roles.push(member.role);
    body.values.push(m);
  });

  console.log('Adding members', body.values);

  await fetch(url, {
    method: 'POST',
    headers,
    body: JSON.stringify(body),
  });
};


const createTeam = async (data) => {
  const owners = await getOwners(data.bearer);
  if (!owners || owners.length === 0) {
    console.error('No owners found');
    return null;
  }

  const { name, description = '' } = data.body;
  if (name) {
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

    // api accepts only 1 member...
    owners.filter(o => o.email.startsWith('admin_ac')).forEach(o => {
      body.members.push({
        '@odata.type': '#microsoft.graph.aadUserConversationMember',
        roles:[
          'owner'
        ],
        'user@odata.bind': `https://graph.microsoft.com/v1.0/users('${o.id}')`
      });
    });

    const res = await fetch(url, {
      method: 'POST',
      headers,
      body: JSON.stringify(body),
    });

    if (res.status === 202) {
      const location = res.headers.get('location');
      const id = location.split('\'')[1];
      console.log('Team created', id);

      //wait 2 seconds... object not found if too fast
      await new Promise(resolve => setTimeout(resolve, 2000));

      await updateTeamPhoto({ bearer: data.bearer, body: { id } });

      const remaining = owners.filter(o => !o.email.startsWith('admin_ac')).map(o => {
        return {
          id: o.id,
          role: 'owner',
        };
      });
      await addMembers(id, remaining, data.bearer);

      return {
        name,
        description,
      };
    }
  }
  return null;
};


export {
  getUser,
  getUserTeams,
  getTeam,
  getTeamById,
  getTeamMembers,
  getAllTeams,
  addRemoveUserToTeams,
  createTeam
}
