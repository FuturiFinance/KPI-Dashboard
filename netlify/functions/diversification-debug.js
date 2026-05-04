// Debug endpoint to inspect runtime environment
// Hit /.netlify/functions/diversification-debug to see what the function can access

exports.handler = async (event, context) => {
  const token = process.env.HUBSPOT_ACCESS_TOKEN;

  const debug = {
    hubspot_token_present: !!token,
    hubspot_token_prefix: token ? token.substring(0, 7) : null,
    hubspot_token_length: token ? token.length : 0,
    all_env_var_keys: Object.keys(process.env).sort(),
    node_version: process.version,
    deploy_id: process.env.DEPLOY_ID || null,
    context_id: process.env.CONTEXT || null,
    site_name: process.env.SITE_NAME || null,
    url: process.env.URL || null,
    timestamp: new Date().toISOString(),
  };

  return {
    statusCode: 200,
    headers: {
      'Content-Type': 'application/json',
      'Cache-Control': 'no-cache, no-store',
    },
    body: JSON.stringify(debug, null, 2),
  };
};
