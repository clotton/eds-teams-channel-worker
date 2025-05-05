const CORS_HEADERS = (env) => {
  return {
    'Access-Control-Allow-Origin': env.CORS_ORIGIN,
    'Access-Control-Allow-Methods': 'GET, POST, PUT, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type, x-api-key',
  };
};

export {
  CORS_HEADERS,
}