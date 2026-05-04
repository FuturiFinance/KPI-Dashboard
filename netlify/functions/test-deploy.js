// Simple test function to verify deployment works
// No dependencies, just returns current timestamp

exports.handler = async (event, context) => {
  return {
    statusCode: 200,
    headers: {
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({
      message: 'Test function deployed successfully',
      timestamp: new Date().toISOString(),
      buildMarker: 'test-deploy-v1',
    }),
  };
};
