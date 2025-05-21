const CORS_HEADERS = (env) => {
  return {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET, POST, PUT, OPTIONS',
    'Access-Control-Allow-Headers': 'Authorization, Content-Type',
  };
};
const TENANT_ID= '36fb8e1d-a891-493f-96d0-3038e4e9291c';

export {
  CORS_HEADERS,
  TENANT_ID
}
