const { API_BASE_URL, AUTH_BASE_URL } = require('zapier-platform-common-microsoft/constants');
const { expect } = require('chai');
const config = require('config');
const zapier = require('zapier-platform-core');
const nock = require('nock');
const App = require('../index');

const appTester = zapier.createAppTester(App);
const authentication = require('../authentication');

describe('authentication', () => {
  it('is of type oauth2', () => expect(authentication.type).to.eql('oauth2'));

  describe('oauth2 config', () => {
    describe('authorizeUrl', () => {
      describe('given inputData and environment are configured correctly', () => {
        let result;
        before(async () => {
          const bundle = {
            inputData: {
              redirect_uri: 'http://zapier.com/',
            },
            environment: {
              CLIENT_ID: config.get('Auth.CLIENT_ID'),
              CLIENT_SECRET: config.get('Auth.CLIENT_SECRET'),
            },
          };

          result = await appTester(App.authentication.oauth2Config.authorizeUrl, bundle);
        });

        it('returns the correct authorize url', () => {
          expect(result).to.eql(
            `${AUTH_BASE_URL}/authorize?scope=offline_access%20user.read%20Calendars.ReadWrite%20Mail.ReadWrite%20Mail.Send%20Contacts.ReadWrite&client_id=${config.get(
              'Auth.CLIENT_ID',
            )}&redirect_uri=http%3A%2F%2Fzapier.com%2F&response_type=code&response_mode=query`,
          );
        });
      });
    });

    describe('getAccessToken', () => {
      describe('given Exchange returns a 200 and the expected response', () => {
        let result;
        before(async () => {
          const bundle = {
            authData: {
              userPrincipalName: 'testzapier@outlook.com',
            },
          };

          nock(AUTH_BASE_URL)
            .post(/token.*/)
            .reply(200, {
              access_token: 'someAccessToken',
              refresh_token: 'someRefreshToken',
            });

          // when we attempt to get the access token
          result = await appTester(App.authentication.oauth2Config.getAccessToken, bundle);
        });

        it('returns the expected access token', () => {
          expect(result.access_token).to.eql('someAccessToken');
        });

        it('returns the expected refresh token', () => {
          expect(result.refresh_token).to.eql('someRefreshToken');
        });

        it('returns the expected scopes', () => {
          expect(result.scopes).to.eql(
            'offline_access user.read Calendars.ReadWrite Mail.ReadWrite Mail.Send Contacts.ReadWrite',
          );
        });

        it('returns the expected userPrincipalName', () => {
          expect(result.userPrincipalName).to.eql('testzapier@outlook.com');
        });
      });

      describe('given Exchange returns a 403 error with some information', () => {
        const invalidAccessTokenResponse = require('./fixtures/responses/invalid_access_token_response'); // eslint-disable-line global-require

        before(() => {
          nock(AUTH_BASE_URL)
            .post(/token.*/)
            .reply(403, invalidAccessTokenResponse);
        });

        it('throws an error with a descriptive message', async () => {
          await expect(
            appTester(App.authentication.oauth2Config.getAccessToken),
          ).to.be.rejectedWith('Unable to fetch access token: Access token is empty.');
        });
      });

      describe('given Exchange returns a 500 error with some weird response', () => {
        before(() => {
          nock(AUTH_BASE_URL)
            .post(/token.*/)
            .reply(500, 'some bizarre string response');
        });

        it('throws an error with a descriptive message', async () => {
          await expect(
            appTester(App.authentication.oauth2Config.getAccessToken),
          ).to.be.rejectedWith(
            'Unable to fetch access token. Error code 500: some bizarre string response',
          );
        });
      });

      describe('given bundle does not have the userPrincipalName stored in it', () => {
        let result;
        let meCall;
        before(async () => {
          const emptyBundle = {};

          nock(AUTH_BASE_URL)
            .post(/token.*/)
            .reply(200, {
              access_token: 'someAccessToken',
              refresh_token: 'someRefreshToken',
            });

          meCall = nock(API_BASE_URL)
            .get('/me')
            .reply(200, { userPrincipalName: 'testzapier@outlook.com' });

          // when we attempt to get the access token
          result = await appTester(App.authentication.oauth2Config.getAccessToken, emptyBundle);
        });

        it('returns the expected access token', () => {
          expect(result.access_token).to.eql('someAccessToken');
        });

        it('returns the expected refresh token', () => {
          expect(result.refresh_token).to.eql('someRefreshToken');
        });

        it('returns the expected scopes', () => {
          expect(result.scopes).to.eql(
            'offline_access user.read Calendars.ReadWrite Mail.ReadWrite Mail.Send Contacts.ReadWrite',
          );
        });

        it('returns the expected userPrincipalName', () => {
          expect(result.userPrincipalName).to.eql('testzapier@outlook.com');
        });

        it('makes an API call to retrieve the userPrincipalName', () => {
          expect(meCall.isDone()).to.eql(true);
        });
      });
    });

    describe('refreshAccessToken', () => {
      describe('given Exchange returns a 200 and the expected response', () => {
        let result;
        before(async () => {
          const bundle = {
            authData: {
              scopes: 'some list of scopes',
              userPrincipalName: 'testzapier@outlook.com',
            },
          };

          nock(AUTH_BASE_URL)
            .post(/token.*/)
            .reply(200, {
              access_token: 'refreshedAccessToken',
              refresh_token: 'newRefreshToken',
            });

          // when we attempt to refresh the access token
          result = await appTester(App.authentication.oauth2Config.refreshAccessToken, bundle);
        });

        it('returns the expected access token', () => {
          expect(result.access_token).to.eql('refreshedAccessToken');
        });

        it('returns the new refresh token', () => {
          expect(result.refresh_token).to.eql('newRefreshToken');
        });

        it('returns the scopes stored in the bundle auth data instead of the scopes constant scopes', () => {
          expect(result.scopes).to.eql('some list of scopes');
        });

        it('returns the expected userPrincipalName', () => {
          expect(result.userPrincipalName).to.eql('testzapier@outlook.com');
        });
      });

      describe('given Exchange returns a 403 error with some information', () => {
        const invalidAccessTokenResponse = require('./fixtures/responses/invalid_access_token_response'); // eslint-disable-line global-require

        before(() => {
          nock(AUTH_BASE_URL)
            .post(/token.*/)
            .reply(403, invalidAccessTokenResponse);
        });

        it('throws an error with a descriptive message', async () => {
          await expect(
            appTester(App.authentication.oauth2Config.refreshAccessToken),
          ).to.be.rejectedWith('Unable to refresh access token: Access token is empty.');
        });
      });

      describe('given Exchange returns a 500 error with some weird response', () => {
        before(() => {
          nock(AUTH_BASE_URL)
            .post(/token.*/)
            .reply(500, 'some bizarre string response');
        });

        it('throws an error with a descriptive message', async () => {
          await expect(
            appTester(App.authentication.oauth2Config.refreshAccessToken),
          ).to.be.rejectedWith(
            'Unable to refresh access token. Error code 500: some bizarre string response',
          );
        });
      });

      describe('given bundle does not have the userPrincipalName stored in it', () => {
        let result;
        let meCall;
        before(async () => {
          const bundle = {
            authData: {
              scopes: 'some list of scopes',
            },
          };

          nock(AUTH_BASE_URL)
            .post(/token.*/)
            .reply(200, {
              access_token: 'refreshedAccessToken',
              refresh_token: 'newRefreshToken',
            });

          meCall = nock(API_BASE_URL)
            .get('/me')
            .reply(200, { userPrincipalName: 'testzapier@outlook.com' });

          // when we attempt to refresh the access token
          result = await appTester(App.authentication.oauth2Config.refreshAccessToken, bundle);
        });

        it('returns the expected access token', () => {
          expect(result.access_token).to.eql('refreshedAccessToken');
        });

        it('returns the new refresh token', () => {
          expect(result.refresh_token).to.eql('newRefreshToken');
        });

        it('returns the scopes stored in the bundle auth data instead of the scopes constant scopes', () => {
          expect(result.scopes).to.eql('some list of scopes');
        });

        it('returns the expected userPrincipalName', () => {
          expect(result.userPrincipalName).to.eql('testzapier@outlook.com');
        });

        it('makes an API call to retrieve the userPrincipalName', () => {
          expect(meCall.isDone()).to.eql(true);
        });
      });
    });

    describe('testAuth', () => {
      describe('given Exchange returns a 200 and the expected response', () => {
        const validTestAuthResponse = require('./fixtures/responses/valid_me_response'); // eslint-disable-line global-require

        let result;
        before(async () => {
          nock(API_BASE_URL)
            .get('/me')
            .reply(200, validTestAuthResponse);

          // when we attempt to test the access token
          result = await appTester(App.authentication.test);
        });

        it('returns the expected access token', () => {
          expect(result.userPrincipalName).to.eql(validTestAuthResponse.userPrincipalName);
        });
      });

      describe('given Exchange returns a 403 error', () => {
        before(() => {
          nock(API_BASE_URL)
            .get('/me')
            .reply(403);
        });

        it('throws an error with a descriptive message', async () => {
          await expect(appTester(App.authentication.test)).to.be.rejectedWith(
            'Something went wrong with the authentication process. Please reconnect your account and try again.',
          );
        });
      });
    });

    describe('given Exchange returns a 500 error with some weird response', () => {
      before(() => {
        nock(API_BASE_URL)
          .get('/me')
          .reply(500, 'some bizarre string response');
      });

      it('throws an error with a descriptive message', async () => {
        await expect(appTester(App.authentication.test)).to.be.rejectedWith(
          'Something went wrong with the authentication process. Please reconnect your account and try again.',
        );
      });
    });
  });
});
