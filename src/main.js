import {CORS_HEADERS, TENANT_ID } from "./constants";
import {
  addRemoveUserToTeams,
  getAllTeams,
  addTeamMembers,
  getTeamMembers,
  getTotalTeamMessages,
  getUserTeams,
  inviteGuest
} from "./api";

const options = async (request, env) => {
  if (request.method === 'OPTIONS') {
    return new Response(null, {
      status: 204,
      headers: CORS_HEADERS(env),
    });
  }
  return new Response('Not found', {
    status: 404,
    headers: CORS_HEADERS(env),
  });
};

const jsonToResponse = async (data, fct, env) => {
  try {
    const json = await fct(data);
    return new Response(JSON.stringify(json || 'Not found'), {
      status: json ? 200 : 404,
      headers: CORS_HEADERS(env),
    });
  } catch (err) {
    return new Response(
      JSON.stringify({ error: err.message || 'Unknown error' }),
      { status: 500, headers: CORS_HEADERS(env) }
    );
  }
};
// Get Microsoft Graph token
async function getGraphToken(env) {
  const res = await fetch(env.AUTH_URL, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      client_id: env.CLIENT_ID,
      client_secret: env.CLIENT_SECRET,
      grant_type: 'client_credentials',
      resource: 'https://graph.microsoft.com',
      scope: 'https://graph.microsoft.com/.default'
    }),
  });

  const json = await res.json();
  return json.access_token;
}
async function getTeamSummary(accessToken, teamId) {
  const url = `https://graph.microsoft.com/v1.0/teams/${teamId}`;

  const response = await fetch(url, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json',
    },
  });

  if (!response.ok) {
    console.warn(
      `Error fetching team ${teamId} summary: ${response.statusText}`);
    return null;
  }

  return await response.json();
}

export default {
  async fetch(request, env) {
    try {
      const url = new URL(request.url);
      const {searchParams } = url;
      const paths = decodeURI(url.pathname).split('/').filter(Boolean);
      let action;
      const data = {};

      if (request.method === 'OPTIONS') {
        return options(request, env);
      }
      if (request.method === 'POST') {
        data.body = await request.json();
      }

      if (paths && paths.length > 0) {
        if (paths[0] === 'teams' || paths[0] === 'users') {
          const prefix = paths[0];
          if (paths.length === 1) {
            action = prefix;
          } else if (paths.length === 2) {
            action = `${prefix}-${paths[1]}`;
          } else if (paths.length === 3) {
            action = `${prefix}-${paths[2]}`;
            data.id = paths[1];
          }
        }
      }

    data.bearer = await getGraphToken(env);

    if (data.bearer) {
      switch (action) {
        case 'teams': {
            data.nameFilter = searchParams.get("nameFilter") || '';
            data. descriptionFilter = searchParams.get("descriptionFilter") || '';
            return jsonToResponse(data, getAllTeams, env);
        }
        case 'teams-invitation': {
          return jsonToResponse({email: data.body.email, name: data.body.name, bearer: data.bearer}, inviteGuest, env);
        }

        case 'teams-summary': {
          const teamIds = data.body.teamIds || [];

          if (!Array.isArray(teamIds) || teamIds.length === 0) {
            return new Response(JSON.stringify({ error: 'No team IDs provided' }), { status: 400, headers: CORS_HEADERS(env) });
          }

          const summaries = await Promise.all(teamIds.map(async (teamId) => {
              const teamSummary = await getTeamSummary(data.bearer, teamId);
              if (!teamSummary) {
                console.warn(`Error fetching team summary for ${teamId}`);
                return null;
              }
              const messageData = await getTotalTeamMessages({id: teamId, bearer: data.bearer});
            return {
              teamId,
              teamName: teamSummary.displayName || '',
              description: teamSummary.description || '',
              created: teamSummary.createdDateTime,
              memberCount: teamSummary.summary.guestsCount + teamSummary.summary.membersCount,
              webUrl: teamSummary.webUrl || '',
              messageCount: messageData.count,
              lastMessage: messageData.latestMessageDate,
            };
          }));

          return new Response(JSON.stringify(summaries.filter(Boolean)), { headers: CORS_HEADERS(env) });
        }
        case 'teams-members': {
          if (request.method === 'GET') {
            return jsonToResponse(data, getTeamMembers, env);
          }
          if (request.method === 'POST') {
            data.tenantId = TENANT_ID;
            return jsonToResponse(data, addTeamMembers, env);
          }
          break;
        }
        case 'users-teams': {
          if (request.method === 'GET') {
            return jsonToResponse(data, getUserTeams, env);
          }
          if (request.method === 'POST') {
            return jsonToResponse(data, addRemoveUserToTeams, env);
          }
          break;
        }
        default:
          return new Response(`Unknown action: ${action}`, {
          status: 404,
          headers: CORS_HEADERS(env),
        });
      }
    } else {
      return new Response('Cannot authenticate to 3rd party', {
        status: 401,
        headers: CORS_HEADERS(env),
      });
    }
  } catch (e) {
      console.error(e);
      return new Response('Oops...', {
        status: 500,
        headers: CORS_HEADERS(env),
      });
    }
  },
};



