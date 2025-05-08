import { CORS_HEADERS } from "./constants";
import { getUserTeams, addRemoveUserToTeams, getAllTeams, getTotalTeamMessages } from "./api";

const jsonToResponse = async (request, data, fct, env) => {
  const json = await fct(data);
  if (json) {
    return new Response(JSON.stringify(json), {
      headers: CORS_HEADERS(env),
    });
  }
  return new Response('Not found', {
    status: 404,
    headers: CORS_HEADERS(env),
  });
}

export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const pathname = url.pathname;
    const searchParams = url.searchParams;
    const nameFilter = searchParams.get("nameFilter") || '';
    const descriptionFilter = searchParams.get("descriptionFilter") || '';

    if (request.method === 'OPTIONS') {
      return new Response(null, { status: 204, headers: CORS_HEADERS(env)});
    }
    let action;
    const data = {
      method: request.method,
    };
    const token = await getGraphToken(env);
    if( pathname === '/teams/addRemoveTeamMember' && request.method === 'POST') {
          const body = await request.json();
          console.log('Received body', body);
          try {
              const emailId = searchParams.get('emailId');
              if (!emailId) {
                  return new Response('Email ID is required', { status: 400 });
              }
              const result = await addRemoveUserToTeams(emailId, body, token);
              console.log('result:', result);
              return new Response(JSON.stringify(result), {
                  headers: CORS_HEADERS(env),
              });
          } catch (err) {
              console.error('Worker error:', err);

              return new Response(
                  JSON.stringify({
                      error: err.message || 'Unknown error',
                      stack: err.stack || '',
                  }),
                  {
                      status: 500,
                      headers:CORS_HEADERS(env),
                  }
              );
          }
      }
      else if (pathname === '/teams/userTeams' && request.method === 'GET') {
          const emailId = searchParams.get('emailId');
          if (!emailId) {
              return new Response('Email ID is required', { status: 400 });
          }
            data.id = emailId;
            data.bearer = token;

             return jsonToResponse(request, data, getUserTeams, env);
      }

      if (pathname === '/teams/allTeams' && request.method === 'GET') {
      try {
        const teams = await getFilteredTeams(token, nameFilter, descriptionFilter);
        console.log(`Fetched ${teams.length} teams`);
        return new Response(JSON.stringify(teams), {
          headers: CORS_HEADERS(env),
        });
      } catch (err) {
        console.error('Worker error:', err);
        return new Response(
            JSON.stringify({
              error: err.message || 'Unknown error',
              stack: err.stack || '',
            }),
            {
              status: 500,
              headers: CORS_HEADERS(env),
            }
        );
      }
    }

    // Handle /teams/summary route
    if (pathname === '/teams/summary' && request.method === 'POST') {
      try {
        const requestBody = await request.json();
        const teamIds = requestBody.teamIds || [];

        if (!Array.isArray(teamIds) || teamIds.length === 0) {
          return new Response(
              JSON.stringify({ error: 'No team IDs provided' }),
              { status: 400,  headers: CORS_HEADERS(env)}
          );
        }

          const teamSummaries = await Promise.all(teamIds.map(async (teamId) => {
              const teamSummary = await getTeamSummary(token, teamId);
              if (!teamSummary) {
                  console.warn(`Error fetching team summary for ${teamId}`);
                  return null;
              }

              const totalMembers = teamSummary.summary.guestsCount + teamSummary.summary.membersCount;
              const messageData = await getTotalTeamMessages(token, teamId);

              return {
                  teamId,
                  teamName: teamSummary.displayName || '',
                  description: teamSummary.description || '',
                  created: teamSummary.createdDateTime,
                  memberCount: totalMembers,
                  webUrl: teamSummary.webUrl || '',
                  messageCount: messageData.messageCount,
                  lastMessage: messageData.latestMessageDate,
              };
          }));

        const validSummaries = teamSummaries.filter(summary => summary !== null);

        return new Response(JSON.stringify(validSummaries), {
            headers: CORS_HEADERS(env),
        });

      } catch (err) {
        console.error('Error in /teams/summary:', err);
        return new Response(
            JSON.stringify({ error: 'Failed to fetch team summaries', details: err.message }),
            { status: 500, headers: CORS_HEADERS(env), }
        );
      }
    }

    return new Response('Not Found', { status: 404 });
  },
};

// Fetch team details from Graph
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
    console.warn(`Error fetching team ${teamId} summary: ${response.statusText}`);
    return null;
  }

  const data = await response.json();
  return data;
}

/// Fetch Teams using only getAllTeams and apply filters
async function getFilteredTeams(accessToken, nameFilter, descriptionFilter) {
  let allTeams = await getAllTeams(accessToken, nameFilter, descriptionFilter);

  // Only keep public teams
  allTeams = allTeams?.filter(o => o.visibility !== 'private');

  const filteredTeams = allTeams.filter(team => {
    const name = team.displayName?.toLowerCase() || '';
    const desc = team.description?.toLowerCase() || '';
    return name.includes(nameFilter.toLowerCase()) &&
        desc.includes(descriptionFilter.toLowerCase());
  });

  return filteredTeams;
}

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




