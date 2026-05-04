// Debug endpoint to inspect runtime environment
// Hit /.netlify/functions/diversification-debug to see what the function can access

exports.handler = async (event, context) => {
  const token = process.env.HUBSPOT_ACCESS_TOKEN;

  const functionsToken = process.env.NETLIFY_FUNCTIONS_TOKEN;
  const siteId = process.env.SITE_ID;

  const debug = {
    hubspot_token_present: !!token,
    hubspot_token_prefix: token ? token.substring(0, 7) : null,
    hubspot_token_length: token ? token.length : 0,
    blobs: {
      site_id_present: !!siteId,
      site_id_value: siteId || null,
      functions_token_present: !!functionsToken,
      functions_token_length: functionsToken ? functionsToken.length : 0,
      functions_token_prefix: functionsToken ? functionsToken.substring(0, 10) : null,
    },
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
