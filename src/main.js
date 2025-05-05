export default {
  async fetch(request, env) {
    const url = new URL(request.url);
    const pathname = url.pathname;
    const searchParams = url.searchParams;
    const nameFilter = searchParams.get("nameFilter") || '';
    const descriptionFilter = searchParams.get("descriptionFilter") || '';

    const token = await getGraphToken(env);

    if (pathname === '/teams/allTeams' && request.method === 'GET') {
      try {
        const teams = await getTeamActivityReport(token, nameFilter, descriptionFilter);
        return new Response(JSON.stringify(teams), {
          headers: {
            'Content-Type': 'application/json',
            'Access-Control-Allow-Origin': '*',
          },
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
              headers: {
                'Content-Type': 'application/json',
                'Access-Control-Allow-Origin': '*',
              },
            }
        );
      }
    }

    return new Response('Not Found', { status: 404 });
  },
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

// Parse CSV string into JSON
function parseCsvToJson(csvString) {
  const lines = csvString.trim().split('\n');
  const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
  return lines.slice(1).map(line => {
    const values = line.split(',').map(v => v.trim().replace(/^"|"$/g, ''));
    return Object.fromEntries(headers.map((h, i) => [h, values[i]]));
  });
}

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

// Throttle concurrent fetches with tracking
let subrequestCount = 0;
const SUBREQUEST_LIMIT = 50;

async function throttledMap(array, limit, asyncFn) {
  const results = [];
  const executing = [];

  for (const item of array) {
    if (subrequestCount >= SUBREQUEST_LIMIT) {
      return new Response("Subrequest limit exceeded", { status: 429 });
    }

    const p = asyncFn(item).then(result => {
      results.push(result);
      subrequestCount++; // Increment count for each subrequest
    });
    executing.push(p);

    if (executing.length >= limit) {
      await Promise.race(executing);
    }
  }

  await Promise.allSettled(executing);
  return results.filter(Boolean);
}

// Fetch Teams activity and enrich with summaries
async function getTeamActivityReport(accessToken, nameFilter = '', descriptionFilter = '') {

  const reportUrl = 'https://graph.microsoft.com/v1.0/reports/getTeamsTeamActivityDetail(period=\'D180\')';

  const response = await fetch(reportUrl, {
    method: 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'text/csv',
    },
  });

  if (!response.ok) {
    throw new Error(`Error fetching report: ${response.statusText}`);
  }

  const contentType = response.headers.get("Content-Type");
  let csvData;

  if (contentType && contentType.includes("application/octet-stream")) {
    const buffer = await response.arrayBuffer();
    csvData = new TextDecoder("utf-8").decode(buffer);
  } else {
    csvData = await response.text();
  }

  const jsonData = parseCsvToJson(csvData);

  const lowerName = nameFilter.trim().toLowerCase();
  const lowerDesc = descriptionFilter.trim().toLowerCase();

  const filteredByName = jsonData.filter(team => {
    const type = team['Team Type']?.toLowerCase() || '';
    const name = (team['Team Name'] || '').toLowerCase();
    const matchesName = !lowerName || name.includes(lowerName);
    return type === 'public' && matchesName;
  });

  // Limit subrequests (e.g., 10 at a time)
  const MAX_CONCURRENT = 10;

  const results = await throttledMap(filteredByName, MAX_CONCURRENT, async team => {
    const teamId = team['Team Id'];
    const teamSummary = await getTeamSummary(accessToken, teamId);

    if (!teamSummary || !teamSummary.summary) return null;

    const description = (teamSummary.description || '').toLowerCase();
    const matchesDescription = !lowerDesc || description.includes(lowerDesc);
    if (!matchesDescription) return null;

    const members = teamSummary.summary.membersCount || 0;
    const guests = teamSummary.summary.guestsCount || 0;
    const total = members + guests;

    return {
      teamId,
      teamName: team['Team Name'] || '',
      description: teamSummary.description || '',
      created: teamSummary.createdDateTime,
      lastActivityDate: team['Last Activity Date'],
      messageCount: team['Channel Messages'],
      activeChannels: team['Active Channels'],
      memberCount: total,
      webUrl: teamSummary.webUrl || '',
    };
  });

  return results;
}
