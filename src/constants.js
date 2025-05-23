const CORS_HEADERS = (env) => {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, OPTIONS',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
  };
};

export {
  CORS_HEADERS
}
