const {
  getAccessToken,
  refreshAccessToken,
  testAuth,
  authorizeUrl,
} = require('zapier-platform-common-microsoft/authentication');

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
