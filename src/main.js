import {CORS_HEADERS} from "./constants";
import {getUserTeams, addRemoveUserToTeams} from "./api";

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
        const teams = await getTeamActivityReport(token, nameFilter, descriptionFilter);
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
              { status: 400, headers: { 'Content-Type': 'application/json' } }
          );
        }

        const teamSummaries = await Promise.all(teamIds.map(async (teamId) => {
          const teamSummary = await getTeamSummary(token, teamId);
          if (!teamSummary) {
            console.warn(`Error fetching team summary for ${teamId}`);
            return null;
          }

          const totalMembers = teamSummary.summary.guestsCount + teamSummary.summary.membersCount;

          return {
            teamId,
            teamName: teamSummary.displayName || '',
            description: teamSummary.description || '',
            created: teamSummary.createdDateTime,
            memberCount: totalMembers,
            webUrl: teamSummary.webUrl || '',
          };
        }));

        // Filter out any null results (in case of errors fetching some summaries)
        const validSummaries = teamSummaries.filter(summary => summary !== null);

        return new Response(JSON.stringify(validSummaries), {
          headers: {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*',
          },
        });

      } catch (err) {
        console.error('Error in /teams/summary:', err);
        return new Response(
            JSON.stringify({ error: 'Failed to fetch team summaries', details: err.message }),
            { status: 500, headers: { 'Content-Type': 'application/json' } }
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

// Fetch Teams activity and enrich with summaries
async function getTeamActivityReport(accessToken, nameFilter, descriptionFilter) {
  const reportUrl = 'https://graph.microsoft.com/v1.0/reports/getTeamsTeamActivityDetail(period=\'D180\')';

  const response = await fetch(reportUrl, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'text/csv',
    },
  });

  if (response.status !== 200) {
    throw new Error(`Error fetching report: ${response.statusText}`);
  }

  let csvData;

  const contentType = response.headers.get("Content-Type");
  if (contentType && contentType.includes("application/octet-stream")) {
    const buffer = await response.arrayBuffer();
    csvData = new TextDecoder("utf-8").decode(buffer);
  } else {
    csvData = await response.text();
  }

  const jsonData = parseCsvToJson(csvData);

  // Apply the filters based on name and description
  const filtered = jsonData.filter(team => {
    const type = team['teamType'] || '';
    const teamName = team['teamName'] || '';

    const matchesNameFilter = teamName.toLowerCase().includes(nameFilter.toLowerCase());

    return type.toLowerCase() === 'public' && matchesNameFilter;
  });

  const allTeams = await getAllTeams(accessToken, nameFilter, descriptionFilter);

  const mergedTeams = mergeTeamsById(filtered, allTeams);

  return mergedTeams;
}

function mergeTeamsById(filtered, allTeams) {
  const allTeamsMap = new Map(allTeams.map(team => [team.id, team]));  // Use 'id' for allTeams

  return filtered
  .filter(team => allTeamsMap.has(team.teamId))  // Use 'teamId' for filtered
  .map(team => ({
    ...team,
    ...allTeamsMap.get(team.teamId),  // Merge using teamId from filtered
  }));
}

// Helper: Convert string to camelCase
function toCamelCase(str) {
  return str
  .toLowerCase()
  .replace(/[^a-z0-9]+(.)/g, (_, chr) => chr.toUpperCase());
}

// Parse CSV string into JSON with camelCase keys
function parseCsvToJson(csvString) {
  const lines = csvString.trim().split('\n');
  const rawHeaders = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
  const headers = rawHeaders.map(toCamelCase);  // Apply camelCase conversion here

  return lines.slice(1).map(line => {
    const values = line.split(',').map(v => v.trim().replace(/^"|"$/g, ''));
    return Object.fromEntries(headers.map((h, i) => [h, values[i]]));
  });
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




