// Diversification Scorecard - HubSpot Data Fetch
// Netlify Function for fetching non-broadcast TopLine Enterprise activity

const HUBSPOT_BASE_URL = 'https://api.hubapi.com';
const PIPELINE_ID = '1265608';

// Stage IDs from Detailed Pipeline
const STAGES = {
  INITIAL_CONTACT: '4906163',
  CONSIDERATION: '4906164',
  PROPOSAL_PRESENTED: '4906173',
  CONTRACT_EXPECTED: '4906175',
  CLOSED_WON: '1608781',
  CLOSED_LOST: '1608782',
};

// Funnel stage groupings (cumulative)
const FUNNEL = {
  CONVERSATIONS: [STAGES.INITIAL_CONTACT, STAGES.CONSIDERATION, STAGES.PROPOSAL_PRESENTED, STAGES.CONTRACT_EXPECTED, STAGES.CLOSED_WON],
  QUALIFIED: [STAGES.CONSIDERATION, STAGES.PROPOSAL_PRESENTED, STAGES.CONTRACT_EXPECTED, STAGES.CLOSED_WON],
  ACTIVE_PROPOSALS: [STAGES.PROPOSAL_PRESENTED, STAGES.CONTRACT_EXPECTED, STAGES.CLOSED_WON],
  CLOSED_WON: [STAGES.CLOSED_WON],
  CLOSED_LOST: [STAGES.CLOSED_LOST],
};

// Targets
const TARGETS = {
  CONVERSATIONS: { july20: 20, yearEnd: 60 },
  QUALIFIED: { july20: 6, yearEnd: 18 },
  ACTIVE_PROPOSALS: { july20: 2, yearEnd: 8 },
  CLOSED_WON: { july20: 1, yearEnd: 3 },
};

// Key dates
const YEAR_START = new Date('2026-01-01T00:00:00-05:00');
const JULY_20 = new Date('2026-07-20T23:59:59-04:00');
const YEAR_END = new Date('2026-12-31T23:59:59-05:00');

function getHeaders() {
  return {
    'Authorization': `Bearer ${process.env.HUBSPOT_ACCESS_TOKEN}`,
    'Content-Type': 'application/json',
  };
}

async function searchDeals(stageIds) {
  const results = [];
  let after = undefined;

  do {
    const body = {
      filterGroups: [{
        filters: [
          { propertyName: 'pipeline', operator: 'EQ', value: PIPELINE_ID },
          { propertyName: 'broadcast_diversification', operator: 'EQ', value: 'true' },
          { propertyName: 'dealstage', operator: 'IN', values: stageIds },
        ]
      }],
      properties: ['dealname', 'dealstage', 'amount', 'closedate', 'createdate', 'hubspot_owner_id'],
      limit: 100,
    };
    if (after) body.after = after;

    const response = await fetch(`${HUBSPOT_BASE_URL}/crm/v3/objects/deals/search`, {
      method: 'POST',
      headers: getHeaders(),
      body: JSON.stringify(body),
    });

    if (!response.ok) {
      const error = await response.text();
      throw new Error(`HubSpot API error: ${response.status} - ${error}`);
    }

    const data = await response.json();
    results.push(...data.results);
    after = data.paging?.next?.after;
  } while (after);

  return results;
}

function calculatePace(target) {
  const now = new Date();
  const daysElapsed = Math.floor((now - YEAR_START) / (1000 * 60 * 60 * 24));
  const daysToJuly20 = Math.floor((JULY_20 - YEAR_START) / (1000 * 60 * 60 * 24));

  // Expected count by today if on pace for July 20 target
  const expectedByNow = (daysElapsed / daysToJuly20) * target;

  return {
    daysElapsed,
    daysToJuly20,
    expectedByNow: Math.round(expectedByNow * 10) / 10,
    pacePercent: (daysElapsed / daysToJuly20) * 100,
  };
}

function getStatus(actual, target) {
  const pace = calculatePace(target);
  const achievementPercent = (actual / pace.expectedByNow) * 100;

  if (achievementPercent >= 100) return 'green';
  if (achievementPercent >= 70) return 'yellow';
  return 'red';
}

exports.handler = async (event, context) => {
  // CORS headers
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
    'Cache-Control': 'public, max-age=3600', // 1 hour cache
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  try {
    // Check for HubSpot token
    if (!process.env.HUBSPOT_ACCESS_TOKEN) {
      return {
        statusCode: 500,
        headers,
        body: JSON.stringify({ error: 'HUBSPOT_ACCESS_TOKEN not configured' }),
      };
    }

    // Fetch all diversification deals (using broadest filter)
    const allStages = [...new Set([...FUNNEL.CONVERSATIONS, ...FUNNEL.CLOSED_LOST])];
    const allDeals = await searchDeals(allStages);

    // Count deals by funnel stage
    const counts = {
      conversations: allDeals.filter(d => FUNNEL.CONVERSATIONS.includes(d.properties.dealstage)).length,
      qualified: allDeals.filter(d => FUNNEL.QUALIFIED.includes(d.properties.dealstage)).length,
      activeProposals: allDeals.filter(d => FUNNEL.ACTIVE_PROPOSALS.includes(d.properties.dealstage)).length,
      closedWon: allDeals.filter(d => FUNNEL.CLOSED_WON.includes(d.properties.dealstage)).length,
      closedLost: allDeals.filter(d => FUNNEL.CLOSED_LOST.includes(d.properties.dealstage)).length,
    };

    // Calculate pace info
    const pace = calculatePace(TARGETS.CONVERSATIONS.july20);

    // Build tiles data
    const tiles = [
      {
        id: 'conversations',
        label: 'Conversations Initiated',
        count: counts.conversations,
        targetJuly20: TARGETS.CONVERSATIONS.july20,
        targetYE: TARGETS.CONVERSATIONS.yearEnd,
        status: getStatus(counts.conversations, TARGETS.CONVERSATIONS.july20),
      },
      {
        id: 'qualified',
        label: 'Qualified Prospects',
        count: counts.qualified,
        targetJuly20: TARGETS.QUALIFIED.july20,
        targetYE: TARGETS.QUALIFIED.yearEnd,
        status: getStatus(counts.qualified, TARGETS.QUALIFIED.july20),
      },
      {
        id: 'activeProposals',
        label: 'Active Proposals',
        count: counts.activeProposals,
        targetJuly20: TARGETS.ACTIVE_PROPOSALS.july20,
        targetYE: TARGETS.ACTIVE_PROPOSALS.yearEnd,
        status: getStatus(counts.activeProposals, TARGETS.ACTIVE_PROPOSALS.july20),
      },
      {
        id: 'closedWon',
        label: 'Closed Pilot or Paid Deal',
        count: counts.closedWon,
        targetJuly20: TARGETS.CLOSED_WON.july20,
        targetYE: TARGETS.CLOSED_WON.yearEnd,
        status: getStatus(counts.closedWon, TARGETS.CLOSED_WON.july20),
      },
      {
        id: 'closedLost',
        label: 'Closed Lost',
        subtitle: 'Learning signal — not a funnel stage',
        count: counts.closedLost,
        isInformational: true,
      },
    ];

    // Calculate conversion rates
    const conversionRates = [
      {
        id: 'conv-to-qualified',
        label: 'Conversation → Qualified',
        numerator: counts.qualified,
        denominator: counts.conversations,
        percent: counts.conversations > 0 ? (counts.qualified / counts.conversations) * 100 : 0,
      },
      {
        id: 'qualified-to-active',
        label: 'Qualified → Active Proposal',
        numerator: counts.activeProposals,
        denominator: counts.qualified,
        percent: counts.qualified > 0 ? (counts.activeProposals / counts.qualified) * 100 : 0,
      },
      {
        id: 'close-rate',
        label: 'Active Proposal → Closed Won',
        subtitle: 'Close Rate (terminal deals only)',
        numerator: counts.closedWon,
        denominator: counts.closedWon + counts.closedLost,
        percent: (counts.closedWon + counts.closedLost) > 0
          ? (counts.closedWon / (counts.closedWon + counts.closedLost)) * 100
          : 0,
      },
    ];

    // Build response
    const response = {
      tiles,
      conversionRates,
      pace: {
        daysElapsed: pace.daysElapsed,
        daysToJuly20: pace.daysToJuly20,
        percentComplete: Math.round(pace.pacePercent * 10) / 10,
      },
      meta: {
        lastUpdated: new Date().toISOString(),
        totalDeals: allDeals.length,
        hubspotViewUrl: `https://app.hubspot.com/contacts/${process.env.HUBSPOT_PORTAL_ID || ''}/objects/0-3/views/all/list?query=broadcast_diversification%3Dtrue`,
      },
    };

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify(response),
    };

  } catch (error) {
    console.error('Diversification fetch error:', error);
    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({
        error: 'Failed to fetch diversification data',
        message: error.message,
      }),
    };
  }
};
