import {CORS_HEADERS} from "./constants";
import {
  getAllTeams,
  createTeam,
  getTeamById,
  addTeamMembers,
  removeTeamMembers,
  getTeamMembers,
  getUserTeams,
  inviteUser,
  handleMessageStatsRequest,
  getTeamMessageStats
} from "./api";

import { requireTurnstileHeader } from "./authentication";
import { isValidReply } from "./utils";


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
    const json = await fct(data, env);
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

async function handleStatsCronJob(env) {
  console.log("handleStatsCronJob started");

  try {
    const bearer = await getGraphToken(env);
    const searchBy = "Message Stats Cron Job";

    const data = { bearer, searchBy, env, nameFilter: '', descriptionFilter: '' };
    const teams = await getAllTeams(data);

    console.log(`Found ${teams.length} teams`);

    for (const team of teams) {
      console.log(`Adding team ${team.displayName} (${team.id}) to stats queue`);
      await env.TEAM_STATS_QUEUE.send({
        teamId: team.id
      });
    }
  } catch (e) {
    console.error("handleStatsCronJob error:", err.stack || err.message);
  }

}

async function processTeamStats(teamId, env) {
  const bearer = await getGraphToken(env);
  let stats = await getTeamMessageStats(teamId, bearer);
  if (!stats) return;

  const key = teamId;
  const newValue = JSON.stringify(stats);
  const existing = await env.TEAMS_KV.get(key);
  console.log(`Getting stats for team ${teamId}: ${existing}`);

  if (existing !== newValue) {
    console.log(`Updating stats for team ${teamId}: ${newValue}`);
    await env.TEAMS_KV.put(key, newValue);
  }
}

async function handleStatsQueue(batch, env, ctx) {
  for (const msg of batch.messages) {
    await processTeamStats(msg.body.teamId, env);
  }
}

export default {
  async scheduled(event, env, ctx) {
    console.log(`[CRON] Triggered at ${new Date().toISOString()}`);
    ctx.waitUntil(handleStatsCronJob(env));
  },

  async queue(batch, env, ctx) {
    await handleStatsQueue(batch, env, ctx);
  },

  async fetch(request, env) {
    try {
      const url = new URL(request.url);
      const { searchParams } = url;
      const paths = decodeURI(url.pathname).split('/').filter(Boolean);
      let action;
      const data = {};

      if (request.method === 'OPTIONS') {
        return options(request, env);
      }

      if (['POST', 'PUT', 'DELETE'].includes(request.method)) {
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
          if (request.method === 'GET') {
            data.searchBy = searchParams.get("searchBy") || '';
            data.nameFilter = searchParams.get("nameFilter") || '';
            data.descriptionFilter = searchParams.get("descriptionFilter") || '';
            data.env = env;
            return jsonToResponse(data, getAllTeams);
          } else if (request.method === 'POST') {
              return jsonToResponse(data, createTeam, env);
            }
          break;
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
            return jsonToResponse(data, getTeamMembers);
          }
          if (request.method === 'POST') {
            const turnstileResponse = await requireTurnstileHeader(request, env);
            if (turnstileResponse) return turnstileResponse;

            return jsonToResponse(data, addTeamMembers, env);
          }
          if (request.method === 'DELETE') {
            const turnstileResponse = await requireTurnstileHeader(request, env);
            if (turnstileResponse) return turnstileResponse;

            return jsonToResponse(data, removeTeamMembers, env);
          }
          break;
        }
        case 'users-teams': {
          if (request.method === 'GET') {
            return jsonToResponse(data, getUserTeams);
          }
          break;
        }
        case 'users-invitation': {
          return jsonToResponse(data, inviteUser);
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
