const { API_BASE_URL } = require('../../constants');

const { expect } = require('chai');
const zapier = require('zapier-platform-core');
const nock = require('nock');
const App = require('../../index');

const appTester = zapier.createAppTester(App);
const getBundle = () => ({
  inputData: {},
});

describe('Microsoft Exchange App', () => {
  it('declares a create contact action', () =>
    expect(App.creates.create_contact).to.exist);

  describe('create contact action', () => {
    describe('given no contact folder is selected and Exchange returns a valid response for valid input', () => {
      const createContactMinFieldsResponse = require('../fixtures/responses/create_contact_min_fields_response'); // eslint-disable-line global-require

      let result;
      let parsedBody;
      let createContactCall;
      before(async () => {
        const bundle = getBundle();
        bundle.inputData.givenName = 'firstName';

        createContactCall = nock(API_BASE_URL)
          .post('/me/contacts', body => {
            parsedBody = body;
            return body;
          })
          .reply(200, createContactMinFieldsResponse);

        // when the user tries to create a contact
        result = await appTester(
          App.creates.create_contact.operation.perform,
          bundle
        );
      });

      it('calls the expected endpoint', () => {
        expect(createContactCall.isDone()).to.eql(true);
      });

      it('calls the APIs with no emails or addresses', () => {
        expect(parsedBody.emailAddresses).to.eql([]);
        expect(parsedBody.businessAddress).to.eql({});
        expect(parsedBody.homeAddress).to.eql({});
        expect(parsedBody.otherAddress).to.eql({});
      });

      it('removes unnecessary attributes from the returned object', () => {
        expect(result).to.not.have.property('@odata.etag');
        expect(result).to.not.have.property('@odata.context');
      });

      it('returns the expected results', () => {
        expect(result.givenName).to.eql(
          createContactMinFieldsResponse.givenName
        );
        expect(result).to.not.have.property('emailAddresses_1');
        expect(result).to.not.have.property('businessPhones_1');
        expect(result).to.not.have.property('homePhones_1');
      });
    });

    describe('given a contact folder is selected and Exchange returns a valid response for valid input', () => {
      const createContactInFolderResponse = require('../fixtures/responses/create_contact_min_fields_response'); // eslint-disable-line global-require

      let result;
      let parsedBody;
      let createContactCall;
      before(async () => {
        const bundle = getBundle();
        bundle.inputData = {
          givenName: 'firstName',
          contactFolderId: 'someContactFolderId',
        };

        createContactCall = nock(API_BASE_URL)
          .post(
            `/me/contactFolders/${bundle.inputData.contactFolderId}/contacts`,
            body => {
              parsedBody = body;
              return body;
            }
          )
          .reply(200, createContactInFolderResponse);

        // when the user tries to create a contact
        result = await appTester(
          App.creates.create_contact.operation.perform,
          bundle
        );
      });

      it('calls the expected endpoint', () => {
        expect(createContactCall.isDone()).to.eql(true);
      });

      it('calls the APIs with no emails or addresses', () => {
        expect(parsedBody.emailAddresses).to.eql([]);
        expect(parsedBody.businessAddress).to.eql({});
        expect(parsedBody.homeAddress).to.eql({});
        expect(parsedBody.otherAddress).to.eql({});
      });

      it('removes unnecessary attributes from the returned object', () => {
        expect(result).to.not.have.property('@odata.etag');
        expect(result).to.not.have.property('@odata.context');
      });

      it('returns the expected results', () => {
        expect(result.givenName).to.eql(
          createContactInFolderResponse.givenName
        );
        expect(result).to.not.have.property('emailAddresses_1');
        expect(result).to.not.have.property('businessPhones_1');
        expect(result).to.not.have.property('homePhones_1');
        expect(result.parentFolderId).to.eql('someContactFolderId');
      });
    });

    describe('given Exchange returns a valid response for valid input of all fields', () => {
      const createContactAllFieldsResponse = require('../fixtures/responses/create_contact_all_fields_response'); // eslint-disable-line global-require

      let result;
      let parsedBody;
      let createContactCall;
      before(async () => {
        const bundle = getBundle();
        bundle.inputData = {
          givenName: 'test',
          surname: 'contact',
          emailAddresses: [
            'test1@test.com',
            'test2@test.com',
            'test3@test.com',
          ],
          businessPhones: ['555-555-5555'],
          homePhones: ['555-555-5554'],
          mobilePhone: '555-555-5553',
          jobTitle: 'test',
          companyName: 'test co',
          department: 'test dept',
          homeAddress: [
            {
              homeAddress_street: '123 home fake st.',
              homeAddress_city: 'austin',
              homeAddress_state: 'tx',
              homeAddress_countryOrRegion: 'us',
              homeAddress_postalCode: '78701',
            },
          ],
          businessAddress: [
            {
              businessAddress_street: '123 business fake st.',
              businessAddress_city: 'austin',
              businessAddress_state: 'tx',
              businessAddress_countryOrRegion: 'us',
              businessAddress_postalCode: '78701',
            },
          ],
          otherAddress: [
            {
              otherAddress_street: '123 other fake st.',
              otherAddress_city: 'austin',
              otherAddress_state: 'tx',
              otherAddress_countryOrRegion: 'us',
              otherAddress_postalCode: '78701',
            },
          ],
        };

        createContactCall = nock(API_BASE_URL)
          .post('/me/contacts', body => {
            parsedBody = body;
            return body;
          })
          .reply(200, createContactAllFieldsResponse);

        // when the user tries to create a contact
        result = await appTester(
          App.creates.create_contact.operation.perform,
          bundle
        );
      });

      it('calls the expected endpoint', () => {
        expect(createContactCall.isDone()).to.eql(true);
      });

      it('calls the APIs with appropriately formatted emails and addresses', () => {
        expect(parsedBody.emailAddresses).to.eql([
          { name: 'test contact', address: 'test1@test.com' },
          { name: 'test contact', address: 'test2@test.com' },
          { name: 'test contact', address: 'test3@test.com' },
        ]);
        expect(parsedBody.businessAddress).to.eql({
          street: '123 business fake st.',
          city: 'austin',
          state: 'tx',
          countryOrRegion: 'us',
          postalCode: '78701',
        });
        expect(parsedBody.homeAddress).to.eql({
          street: '123 home fake st.',
          city: 'austin',
          state: 'tx',
          countryOrRegion: 'us',
          postalCode: '78701',
        });
        expect(parsedBody.otherAddress).to.eql({
          street: '123 other fake st.',
          city: 'austin',
          state: 'tx',
          countryOrRegion: 'us',
          postalCode: '78701',
        });
      });

      it('removes unnecessary attributes from the returned object', () => {
        expect(result).to.not.have.property('@odata.etag');
        expect(result).to.not.have.property('@odata.context');
      });

      it('returns the expected results', () => {
        expect(result.givenName).to.eql(
          createContactAllFieldsResponse.givenName
        );
        expect(result.emailAddresses_1).to.eql('test1@test.com');
        expect(result.emailAddresses_2).to.eql('test2@test.com');
        expect(result.emailAddresses_3).to.eql('test3@test.com');
        expect(result.businessPhones_1).to.eql('555-555-5555');
        expect(result.homePhones_1).to.eql('555-555-5554');
      });
    });

    describe('given Exchange returns a descriptive error', () => {
      const descriptiveError = require('../fixtures/responses/invalid_access_token_response'); // eslint-disable-line global-require

      before(async () => {
        nock(API_BASE_URL)
          .post('/me/contacts')
          .reply(400, descriptiveError);
      });

      it('returns a descriptive error message', async () => {
        await expect(
          appTester(App.creates.create_contact.operation.perform)
        ).to.be.rejectedWith(
          'Unable to create a contact: Access token is empty.'
        );
      });
    });

    describe('given Exchange returns an Access Denied error', () => {
      const accessDeniedResponse = require('../fixtures/responses/error_access_denied_response'); // eslint-disable-line global-require

      before(async () => {
        nock(API_BASE_URL)
          .post('/me/contacts')
          .reply(403, accessDeniedResponse);
      });

      it('returns a descriptive error message', async () => {
        await expect(
          appTester(App.creates.create_contact.operation.perform)
        ).to.be.rejectedWith(
          'Unable to create a contact: This feature requires new permissions from your Exchange account. Please reconnect your account to take advantage of it.'
        );
      });
    });

    describe('given Exchange returns an unusual error (500)', () => {
      before(() => {
        nock(API_BASE_URL)
          .post('/me/contacts')
          .reply(500, 'some internal server error');
      });

      it('returns a descriptive error message', async () => {
        await expect(
          appTester(App.creates.create_contact.operation.perform)
        ).to.be.rejectedWith(
          'Unable to create a contact. Error code 500: some internal server error'
        );
      });
    });
  });
});
