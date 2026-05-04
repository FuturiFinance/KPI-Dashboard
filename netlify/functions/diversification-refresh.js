// Scheduled function to refresh diversification data cache
// Runs daily at 6am ET (11:00 UTC)

const { handler: diversificationHandler } = require('./diversification');

exports.handler = async (event, context) => {
  console.log('Scheduled diversification data refresh starting...');

  try {
    // Call the main diversification handler to warm the cache
    const result = await diversificationHandler(
      { httpMethod: 'GET' },
      context
    );

    console.log('Diversification data refresh completed:', {
      statusCode: result.statusCode,
      timestamp: new Date().toISOString(),
    });

    return {
      statusCode: 200,
      body: JSON.stringify({
        message: 'Diversification cache refreshed',
        timestamp: new Date().toISOString(),
      }),
    };
  } catch (error) {
    console.error('Diversification refresh error:', error);
    return {
      statusCode: 500,
      body: JSON.stringify({
        error: 'Refresh failed',
        message: error.message,
      }),
    };
  }
};
