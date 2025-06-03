import {CORS_HEADERS} from "./constants";
import {
  getAllTeams,
  getTeamById,
  addTeamMembers,
  getTeamMembers,
  getUserTeams,
  inviteUser,
  handleMessageStatsRequest,
  getTeamMessageStats
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

async function processTeamStats({ teamId }, env) {
  const bearer = await getGraphToken(env);
  const stats = await getTeamMessageStats(teamId, bearer);
  await env.TEAM_KV.put(`stats:${teamId}`, JSON.stringify(stats));
}

async function generateJobs(env) {
  const bearer = await getGraphToken(env);
  const data = { };
  data.bearer = bearer;
  data.nameFilter = 'aem-';
  data.descriptionFilter = 'Edge Delivery';

  const teams = await getAllTeams(data);
  return teams.map(team => ({
    body: JSON.stringify({ teamId: team.id })
  }));
}

export default {
  async scheduled(event, env, ctx) {
    const messages = await generateJobs(env);
    ctx.waitUntil(env.TEAM_STATS_QUEUE.sendBatch(messages));
  },

  async queue(batch, env, ctx) {
    for (const msg of batch.messages) {
      ctx.waitUntil(processTeamStats(msg.body, env));
    }
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
