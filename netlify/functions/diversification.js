// Diversification Scorecard - reads from stored snapshot
// Snapshot is updated weekly (Monday 6am ET) or manually via refresh endpoint

const { getStore, connectLambda } = require('@netlify/blobs');

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
    // Initialize Lambda context for Netlify Blobs
    connectLambda(event);

    // Read snapshot from Netlify Blobs
    const store = getStore(STORE_NAME);
    const data = await store.get(SNAPSHOT_KEY, { type: 'json' });

    if (!data) {
      return {
        statusCode: 404,
        headers,
        body: JSON.stringify({
          error: 'No snapshot available',
          message: 'No diversification data has been captured yet. Click "Refresh Data" to fetch from HubSpot.',
          errorType: 'NO_SNAPSHOT',
        }),
      };
    }

    return {
      statusCode: 200,
      headers,
      body: JSON.stringify(data),
    };

  } catch (error) {
    console.error('Error reading snapshot:', error);

    // Handle case where blob doesn't exist
    if (error.message && error.message.includes('not found')) {
      return {
        statusCode: 404,
        headers,
        body: JSON.stringify({
          error: 'No snapshot available',
          message: 'No diversification data has been captured yet. Click "Refresh Data" to fetch from HubSpot.',
          errorType: 'NO_SNAPSHOT',
        }),
      };
    }

    return {
      statusCode: 500,
      headers,
      body: JSON.stringify({
        error: 'Failed to read snapshot',
        message: error.message,
        errorType: 'STORAGE_ERROR',
      }),
    };
  }
};
