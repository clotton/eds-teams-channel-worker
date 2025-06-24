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
  getTeamMessageStats, getMessagesLast30Days
} from "./api";

import { requireTurnstileHeader } from "./authentication";
import { isQuestion } from "./utils";


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

async function handleMessageStatsCronJob(env) {
  const bearer = await getGraphToken(env);
  const searchBy = "Message Stats Cron Job";

  const data = { bearer, searchBy, env, nameFilter: '', descriptionFilter: '' };
  const teams = await getAllTeams(data);

  for (const team of teams) {
    await env.TEAM_STATS_QUEUE.send({
      teamId: team.id
    });
  }
}

async function handleAnalyticsCronJob(env) {
  const bearer = await getGraphToken(env);
  const searchBy = "Monthly Analytics Cron Job";

  const data = { bearer, searchBy, env, nameFilter: '', descriptionFilter: '' };
  const teams = await getAllTeams(data);

  let createdLast30DaysCount = 0;
  let questionsLast30DaysCount = 0;

  for (const team of teams) {
    const teamStats = await getTeamById({ id: team.id, bearer });
    console.log(`Processing team: ${team.displayName} (${team.id}) with createdDateTime of ${teamStats.createdDateTime}` );
    if (teamStats.createdDateTime && new Date(teamStats.createdDateTime) > new Date(Date.now() - 30 * 24 * 60 * 60 * 1000)) {
      createdLast30DaysCount++;
    }

    const teamMessage30Days = await getMessagesLast30Days(team.id, bearer);
    // Then to count questions:
    const allMessages = teamMessage30Days.messages || [];
    const questionMessages = allMessages.filter(msg => isQuestion(msg.body?.content));
    const questionCount = questionMessages.length;

    questionsLast30DaysCount += questionCount;
  }

  await env.TEAMS_ANALYTICS_QUEUE.send({
    created_30_days: createdLast30DaysCount,
    questions_30_days: questionsLast30DaysCount
  });
}

async function processTeamMessageStats(teamId, env) {
  const bearer = await getGraphToken(env);
  const stats = await getTeamMessageStats(teamId, bearer);
  if (!stats) return;

  const key = teamId;
  const newValue = JSON.stringify(stats);
  const existing = await env.TEAMS_KV.get(key);

  if (existing !== newValue) {
    await env.TEAMS_KV.put(key, newValue);
  }
}

async function processTeamAnalytics(data, env) {
  await env.TEAMS_KV.put("created_last_30_days", data.created_30_days.toString());
  await env.TEAMS_KV.put("questions_last_30_days", data.questions_30_days.toString());

}

async function handleStatsQueue(batch, env, ctx) {
  const chunkSize = 1;
  for (let i = 0; i < batch.messages.length; i += chunkSize) {
    const chunk = batch.messages.slice(i, i + chunkSize);
    ctx.waitUntil(Promise.all(chunk.map(msg =>
        processTeamMessageStats(msg.body.teamId, env)
    )));
  }
}

async function handleAnalyticsQueue(batch, env, ctx) {
  const chunkSize = 1;
  for (let i = 0; i < batch.messages.length; i += chunkSize) {
    const chunk = batch.messages.slice(i, i + chunkSize);
    ctx.waitUntil(Promise.all(chunk.map(msg =>
        processTeamAnalytics(msg.body, env)
    )));
  }
}

export default {
  async scheduled(event, env, ctx) {
    if (event.cron === "0 */12 * * *") {
      ctx.waitUntil(handleMessageStatsCronJob(env));
    } else if (event.cron === "*/10 * * * *") {
      ctx.waitUntil(handleAnalyticsCronJob(env));
    }
  },

  async queue(batch, env, ctx) {
    switch (batch.queue) {
      case "team-stats-queue":
        await handleStatsQueue(batch, env, ctx);
        break;
      case "teams-analytics-queue":
        await handleAnalyticsQueue(batch, env, ctx);
        break;
    }
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
