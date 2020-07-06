/**
 * Authentication in a Zapier integration can be done in one of multiple ways. The recommended
 * authentication scheme is OAuth2, which is what we're using here.
 *
 * More details here: https://platform.zapier.com/cli_docs/docs#oauth2
 */

const { getUserPrincipalName } = require('./utils');
const { API_BASE_URL, AUTH_BASE_URL } = require('./constants');

const config = require('config');

// offline_access - needed for refresh token
// user.read - needed for testing auth
const scopes = 'offline_access user.read Contacts.ReadWrite';

async function getAccessToken(z, bundle) {
  const response = await z.request(`${AUTH_BASE_URL}/token`, {
    method: 'POST',
    body: {
      grant_type: 'authorization_code',
      client_id: process.env.CLIENT_ID || '1234',
      client_secret: process.env.CLIENT_SECRET || 'abcd',
      redirect_uri: `${bundle.inputData.redirect_uri}`,
      code: bundle.inputData.code,
      scope: scopes,
    },
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    prefixErrorMessageWith: 'Unable to fetch access token',
  });

  const { access_token, refresh_token } = z.JSON.parse(response.content);

  // Need to have the access token set to make an API call in the method below this
  bundle.authData.access_token = access_token; // eslint-disable-line no-param-reassign
  const userPrincipalName = await getUserPrincipalName(z, bundle);

  // We need to save the scopes so that we can safely migrate users to newer
  // versions with newer scopes without breaking their Zaps. Both the get and
  // refresh calls require us to specify which scopes we want and if we didn't
  // save the scopes that were being used, all Zaps would break if we changed
  // the higher level required scopes.
  //
  // If users want to use new features if we add new scopes at the top of this
  // file, then they will need to reconnect their account. We can provide a
  // better error message at that time if needed.
  //
  // We also want to save the userPrincipalName as that is needed for filtering
  // on certain APIs and it doesn't make sense to re-request that each time.
  return {
    access_token,
    refresh_token,
    scopes,
    userPrincipalName,
  };
}

async function refreshAccessToken(z, bundle) {
  const response = await z.request(`${AUTH_BASE_URL}/token`, {
    method: 'POST',
    body: {
      grant_type: 'refresh_token',
      client_id: process.env.CLIENT_ID || '1234',
      client_secret: process.env.CLIENT_SECRET || 'abcd',
      redirect_uri: `${bundle.inputData.redirect_uri}`,
      refresh_token: `${bundle.authData.refresh_token}`,
      scope: `${bundle.authData.scopes}`,
    },
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    },
    prefixErrorMessageWith: 'Unable to refresh access token',
  });

  const { access_token, refresh_token } = z.JSON.parse(response.content);

  // Need to have the access token set to make an API call in the method below this
  bundle.authData.access_token = access_token; // eslint-disable-line no-param-reassign
  const userPrincipalName = await getUserPrincipalName(z, bundle);

  return {
    access_token,
    refresh_token,
    scopes: bundle.authData.scopes,
    userPrincipalName,
  };
}

async function testAuth(z) {
  const response = await z.request({
    method: 'GET',
    url: `${API_BASE_URL}/me`,
    disableMiddlewareErrorChecking: true,
  });

  if (response.status !== 200) {
    throw new Error(
      'Something went wrong with the authentication process. ' +
        'Please reconnect your account and try again.'
    );
  }
  return z.JSON.parse(response.content);
}

async function authorizeUrl(z, bundle) {
  let url = `${AUTH_BASE_URL}/authorize`;

  const urlParts = [
    `scope=${encodeURIComponent(scopes)}`,
    `client_id=${process.env.CLIENT_ID || '1234'}`,
    `redirect_uriprocess.env.CLIENT_SECRET || 'abcd'ta.redirect_uri)}`,
    `response_type=code`,
    `response_mode=query`,
  ];

  return `${url}?${urlParts.join('&')}`;
}

module.exports = {
  type: 'oauth2',
  oauth2Config: {
    authorizeUrl,
    getAccessToken,
    refreshAccessToken,
    autoRefresh: true,
  },
  test: testAuth,
  connectionLabel: '{{userPrincipalName}}',
};
