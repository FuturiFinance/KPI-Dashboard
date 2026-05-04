// Refresh endpoint: fetches from HubSpot and stores snapshot in Netlify Blobs
// Can be called manually or by scheduled function

const { getStore } = require('@netlify/blobs');
const { fetchDiversificationData } = require('./lib/hubspot-diversification');

const STORE_NAME = 'diversification';
const SNAPSHOT_KEY = 'current-snapshot';

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
    console.log('Refreshing diversification data from HubSpot...');

    // Fetch fresh data from HubSpot
    const data = await fetchDiversificationData();

    // Add refresh metadata
    data.meta.refreshedAt = new Date().toISOString();
    data.meta.refreshSource = event.httpMethod === 'POST' ? 'manual' : 'scheduled';

    // Store in Netlify Blobs
    const store = getStore(STORE_NAME);
    await store.setJSON(SNAPSHOT_KEY, data);

    console.log('Snapshot stored successfully:', {
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
    console.error('Refresh failed:', error);

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
