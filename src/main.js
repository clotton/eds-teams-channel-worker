// Cloudflare Worker backend for Microsoft Teams with rate limiting handling

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
          headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
        });
      } catch (err) {
        return new Response(JSON.stringify({ error: 'Failed to fetch all teams' }), {
          status: 500,
          headers: { 'Content-Type': 'application/json', 'Access-Control-Allow-Origin': '*' },
        });
      }
    }

    return new Response('Not Found', { status: 404 });
  },
};

async function getGraphToken(env) {
  const res = await fetch(env.AUTH_URL, {
    method: "POST",
    headers: {"Content-Type": "application/x-www-form-urlencoded"},
    body: new URLSearchParams({
      client_id: env.CLIENT_ID,
      client_secret: env.CLIENT_SECRET,
      grant_type: 'client_credentials',
      resource: 'https://graph.microsoft.com',
      scope: 'https://graphs.microsoft.com/.default'
    })
  });

  const json = await res.json();
  return json.access_token;
}

// Helper: parse CSV to JSON
function parseCsvToJson(csvString) {
  const lines = csvString.trim().split('\n');
  const headers = lines[0].split(',').map(h => h.trim().replace(/^"|"$/g, ''));
  return lines.slice(1).map(line => {
    const values = line.split(',').map(v => v.trim().replace(/^"|"$/g, ''));
    return Object.fromEntries(headers.map((h, i) => [h, values[i]]));
  });
}

// Main function: fetch, filter, and format Teams activity report
async function getTeamActivityReport(accessToken, nameFilter = '', descriptionFilter = '') {
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

  // Check if Content-Type is application/octet-stream and handle accordingly
  const contentType = response.headers.get("Content-Type");
  if (contentType && contentType.includes("application/octet-stream")) {
    // Use .arrayBuffer() for binary data
    const buffer = await response.arrayBuffer();
    csvData = new TextDecoder("utf-8").decode(buffer);
    // Now parse the CSV data to JSON
  } else {
    // If it's text, we can directly use .text()
    csvData = await response.text();
  }

  const jsonData = parseCsvToJson(csvData);

  // Filter: name/description matches + team type is 'Public'
  const filtered = jsonData.filter(team => {
    const name = team['Team Name'] || team['Team Display Name'] || '';
    const description = team['Team Description'] || '';
    const type = team['Team Type'] || '';
    return name.toLowerCase().includes(nameFilter.toLowerCase()) &&
        type.toLowerCase() === 'public';
  });

  // Simplify: keep only selected fields
  const simplified = filtered.map(team => ({
    teamId: team['Team Id'],
    teamName: team['Team Name'] || team['Team Display Name'],
    lastActivityDate: team['Last Activity Date'],
    channelMessages: team['Channel Messages'] || team['Number of Channel Messages'],
    activeChannels: team['Active Channels'],
    memberCount: team['Active Users'] || team['Number of team members'],
  }));

  return simplified;
}




