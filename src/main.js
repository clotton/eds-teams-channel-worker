import {CORS_HEADERS} from "./constants";
import {
  getAllTeams,
  getTeamById,
  addTeamMembers,
  getTeamMembers,
  getUserTeams,
  inviteUser,
  handleMessageStatsRequest
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

async function fetchAndCacheAllTeamStats(env) {
  try {
    const bearer = await getGraphToken(env);
    if (!bearer) {
      console.error('Failed to get Graph token');
      return;
    }

    const nameFilter = "aem-";
    const descriptionFilter = "Edge Delivery";

    const teams = await getAllTeams({ bearer, nameFilter, descriptionFilter });
    if (!teams || teams.length === 0) {
      console.warn('No teams found');
      return;
    }

    await Promise.all(teams.map(async (team) => {
      const teamDetails = await getTeamById({ id: team.id, bearer });

      if (!teamDetails) {
        console.warn(`Error fetching team ${team.id} details`);
        return null;
      }

      let teamSummary = {
        teamId: team.id,
        teamName: teamDetails.displayName || '',
        description: teamDetails.description || '',
        created: teamDetails.createdDateTime,
        memberCount: teamDetails.summary.guestsCount + teamDetails.summary.membersCount,
        webUrl: teamDetails.webUrl || '',
      };

      // Cache the summaries in KV or any other storage
      await env.TEAMS_KV.put(team.id, JSON.stringify(teamSummary));

      return {
        teamId: team.id,
        teamName: teamDetails.displayName || '',
        description: teamDetails.description || '',
        created: teamDetails.createdDateTime,
        memberCount: teamDetails.summary.guestsCount + teamDetails.summary.membersCount,
        webUrl: teamDetails.webUrl || '',
      };
    }));


  } catch (error) {
    console.error('Error fetching and caching all team stats:', error);
  }
}


export default {
  async scheduled(event, env, ctx) {
    ctx.waitUntil(fetchAndCacheAllTeamStats(env));
  },
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
            data.descriptionFilter = searchParams.get("descriptionFilter") || '';
            return jsonToResponse(data, getAllTeams, env);
        }
        case 'teams-summary': {
          const teamIds = data.body.teamIds || [];

          if (!Array.isArray(teamIds) || teamIds.length === 0) {
            return new Response(JSON.stringify({ error: 'No team IDs provided' }), { status: 400, headers: CORS_HEADERS(env) });
          }

          const summaries = await Promise.all(teamIds.map(async (teamId) => {
            const team = await getTeamById({ id: teamId, bearer: data.bearer });
            if (!team) {
              console.warn(`Error fetching team ${teamId} summary`);
              return null;
            }
            return {
              teamId,
              teamName: team.displayName || '',
              description: team.description || '',
              created: team.createdDateTime,
              memberCount: team.summary.guestsCount + team.summary.membersCount,
              webUrl: team.webUrl || '',
            };
          }));

          return new Response(JSON.stringify(summaries.filter(Boolean)), { headers: CORS_HEADERS(env) });
        }
        case 'teams-messages': {
          if (request.method === 'POST') {
            return jsonToResponse(data, handleMessageStatsRequest, env);
          }
          break;
        }
        case 'teams-members': {
          if (request.method === 'GET') {
            return jsonToResponse(data, getTeamMembers, env);
          }
          if (request.method === 'POST') {
            data.env = env;
            return jsonToResponse(data, addTeamMembers, env);
          }
          break;
        }
        case 'users-teams': {
          if (request.method === 'GET') {
            return jsonToResponse(data, getUserTeams, env);
          }
          break;
        }
        case 'users-invitation': {
          return jsonToResponse(data, inviteUser, env);
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
