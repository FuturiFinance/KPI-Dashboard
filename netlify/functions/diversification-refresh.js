// Manual refresh endpoint - HTTP accessible
// Fetches from HubSpot and stores snapshot in Netlify Blobs

const { getStore, connectLambda } = require('@netlify/blobs');

const STORE_NAME = 'diversification';
const SNAPSHOT_KEY = 'current-snapshot';

const HUBSPOT_BASE_URL = 'https://api.hubapi.com';
const PIPELINE_ID = '1265608';

const STAGES = {
  INITIAL_CONTACT: '4906163',
  CONSIDERATION: '4906164',
  PROPOSAL_PRESENTED: '4906173',
  CONTRACT_EXPECTED: '4906175',
  CLOSED_WON: '1608781',
  CLOSED_LOST: '1608782',
};

// Human-readable stage names for display
const STAGE_NAMES = {
  '4906163': 'Initial Contact',
  '4906164': 'Consideration',
  '4906173': 'Proposal Presented',
  '4906175': 'Contract Expected',
  '1608781': 'Closed Won',
  '1608782': 'Closed Lost',
};

const FUNNEL = {
  CONVERSATIONS: [STAGES.INITIAL_CONTACT, STAGES.CONSIDERATION, STAGES.PROPOSAL_PRESENTED, STAGES.CONTRACT_EXPECTED, STAGES.CLOSED_WON],
  QUALIFIED: [STAGES.CONSIDERATION, STAGES.PROPOSAL_PRESENTED, STAGES.CONTRACT_EXPECTED, STAGES.CLOSED_WON],
  ACTIVE_PROPOSALS: [STAGES.PROPOSAL_PRESENTED, STAGES.CONTRACT_EXPECTED, STAGES.CLOSED_WON],
  CLOSED_WON: [STAGES.CLOSED_WON],
  CLOSED_LOST: [STAGES.CLOSED_LOST],
};

const TARGETS = {
  CONVERSATIONS: { july20: 20, yearEnd: 60 },
  QUALIFIED: { july20: 6, yearEnd: 18 },
  ACTIVE_PROPOSALS: { july20: 2, yearEnd: 8 },
  CLOSED_WON: { july20: 1, yearEnd: 3 },
};

// Mandate start date - May 4, 2026 (fixed, not dynamic)
const MANDATE_START = new Date('2026-05-04T00:00:00-04:00');
const JULY_20 = new Date('2026-07-20T23:59:59-04:00');
const YEAR_END = new Date('2026-12-31T23:59:59-05:00');

// Window lengths in days
const DAYS_TO_JULY_20 = 77;  // May 4 to July 20
const DAYS_TO_YEAR_END = 241;  // May 4 to Dec 31

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
      const errorBody = await response.text();
      throw new Error(`HubSpot API error ${response.status}: ${errorBody}`);
    }

    const data = await response.json();
    results.push(...data.results);
    after = data.paging?.next?.after;
  } while (after);

  return results;
}

function calculatePace(julyTarget) {
  const now = new Date();
  const daysElapsed = Math.max(0, Math.floor((now - MANDATE_START) / (1000 * 60 * 60 * 24)));
  const expectedByNow = (daysElapsed / DAYS_TO_JULY_20) * julyTarget;

  return {
    daysElapsed,
    totalDays: DAYS_TO_JULY_20,
    expectedByNow: Math.round(expectedByNow * 10) / 10,
    pacePercent: Math.round((daysElapsed / DAYS_TO_JULY_20) * 1000) / 10,
  };
}

function getStatus(actual, julyTarget) {
  const pace = calculatePace(julyTarget);

  // On day 0 (or before mandate start), expected is 0, so any count is green
  if (pace.expectedByNow <= 0) {
    return 'green';
  }

  const achievementPercent = (actual / pace.expectedByNow) * 100;

  if (achievementPercent >= 100) return 'green';
  if (achievementPercent >= 70) return 'yellow';
  return 'red';
}

// Format a deal for frontend display
function formatDeal(deal) {
  const props = deal.properties || {};
  const stageId = props.dealstage;

  return {
    id: deal.id,
    name: props.dealname || 'Unnamed Deal',
    stage: STAGE_NAMES[stageId] || stageId,
    stageId: stageId,
    amount: props.amount ? parseFloat(props.amount) : null,
    closeDate: props.closedate || null,
    createDate: props.createdate || null,
    hubspotUrl: `https://app.hubspot.com/contacts/${process.env.HUBSPOT_PORTAL_ID || '6154760'}/deal/${deal.id}`,
  };
}

async function fetchDiversificationData() {
  if (!process.env.HUBSPOT_ACCESS_TOKEN) {
    throw new Error('HUBSPOT_ACCESS_TOKEN not configured');
  }

  const allStages = [...new Set([...FUNNEL.CONVERSATIONS, ...FUNNEL.CLOSED_LOST])];
  const allDeals = await searchDeals(allStages);

  // Group deals by funnel category
  const dealsByCategory = {
    conversations: allDeals.filter(d => FUNNEL.CONVERSATIONS.includes(d.properties.dealstage)),
    qualified: allDeals.filter(d => FUNNEL.QUALIFIED.includes(d.properties.dealstage)),
    activeProposals: allDeals.filter(d => FUNNEL.ACTIVE_PROPOSALS.includes(d.properties.dealstage)),
    closedWon: allDeals.filter(d => FUNNEL.CLOSED_WON.includes(d.properties.dealstage)),
    closedLost: allDeals.filter(d => FUNNEL.CLOSED_LOST.includes(d.properties.dealstage)),
  };

  const counts = {
    conversations: dealsByCategory.conversations.length,
    qualified: dealsByCategory.qualified.length,
    activeProposals: dealsByCategory.activeProposals.length,
    closedWon: dealsByCategory.closedWon.length,
    closedLost: dealsByCategory.closedLost.length,
  };

  // Format deals for frontend display
  const deals = {
    conversations: dealsByCategory.conversations.map(formatDeal),
    qualified: dealsByCategory.qualified.map(formatDeal),
    activeProposals: dealsByCategory.activeProposals.map(formatDeal),
    closedWon: dealsByCategory.closedWon.map(formatDeal),
    closedLost: dealsByCategory.closedLost.map(formatDeal),
  };

  const pace = calculatePace(TARGETS.CONVERSATIONS.july20);

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

  return {
    tiles,
    conversionRates,
    deals,
    pace: {
      daysElapsed: pace.daysElapsed,
      totalDays: pace.totalDays,
      percentComplete: pace.pacePercent,
    },
    meta: {
      lastUpdated: new Date().toISOString(),
      totalDeals: allDeals.length,
      hubspotViewUrl: `https://app.hubspot.com/contacts/${process.env.HUBSPOT_PORTAL_ID || '6154760'}/objects/0-3/views/all/list?query=broadcast_diversification%3Dtrue`,
    },
  };
}

exports.handler = async (event, context) => {
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Headers': 'Content-Type',
    'Content-Type': 'application/json',
    'Cache-Control': 'no-cache, no-store, must-revalidate',
  };

  if (event.httpMethod === 'OPTIONS') {
    return { statusCode: 200, headers, body: '' };
  }

  try {
    console.log('Manual refresh triggered...');

    // Initialize Lambda context for Netlify Blobs
    connectLambda(event);

    // Fetch fresh data from HubSpot
    const data = await fetchDiversificationData();

    // Add refresh metadata
    data.meta.refreshedAt = new Date().toISOString();
    data.meta.refreshSource = 'manual';

    // Store in Netlify Blobs
    const store = getStore(STORE_NAME);
    await store.setJSON(SNAPSHOT_KEY, data);

    console.log('Manual refresh completed:', {
      totalDeals: data.meta.totalDeals,
      refreshedAt: data.meta.refreshedAt,
    });

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify({
        success: true,
        message: 'Diversification data refreshed',
        refreshedAt: data.meta.refreshedAt,
        totalDeals: data.meta.totalDeals,
      }),
    };

  } catch (error) {
    console.error('Manual refresh failed:', error);

    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({
        success: false,
        error: 'Refresh failed',
        message: error.message,
      }),
    };
  }
};
