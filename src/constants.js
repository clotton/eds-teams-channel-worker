const CORS_HEADERS = (env) => {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, DELETE, OPTIONS',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type, x-turnstile-token',
  };
};

export {
  CORS_HEADERS
}
